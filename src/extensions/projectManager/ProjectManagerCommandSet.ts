import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/sites";
import "@pnp/sp/user-custom-actions";
import { IUserCustomActionInfo } from '@pnp/sp/user-custom-actions';
import TemplateLoaderDialog from './TemplateLoaderDialog';
import { TemplateSettingsDialog, ITemplateSettings } from './TemplateSettingsDialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProjectManagerCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
  templateSettings?: ITemplateSettings;
  isConfigured?: boolean;
}

const LOG_SOURCE: string = 'ProjectManagerCommandSet';
const EXTENSION_ID: string = '4931b98a-cfa0-423a-95e3-d76ac877c1ae'; // From manifest.json

// Custom interface to include ClientSideComponentId property
interface IAppCustomizerInfo extends IUserCustomActionInfo {
  ClientSideComponentId: string;
  ClientSideComponentProperties: string;
}

export interface IFileItem {
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;
  TimeLastModified: string;
  Length: number;
  UniqueId: string;
  Program?: string;
}

export interface IFolderItem {
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;
  TimeLastModified: string;
  ItemCount: number;
  UniqueId: string;
  Program?: string;
}

export default class ProjectManagerCommandSet extends BaseListViewCommandSet<IProjectManagerCommandSetProperties> {
  private _sp: SPFI;
  private _checkInterval: number | undefined;
  private _isLoadingTemplate: boolean = false;
  private _templateSettings: ITemplateSettings | null = null;
  private _lastCheckedFolder: string = '';
  private _folderCheckCache: { [key: string]: { isEmpty: boolean; checkedAt: number } } = {};

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ProjectManagerCommandSet');

    // Initialize PnP
    this._sp = spfi().using(SPFx(this.context));

    // Load saved template settings
    await this._loadTemplateSettings();

    // Only start monitoring if settings are configured and in automatic mode
    if (this._templateSettings) {
      if (this._templateSettings.copyMode === 'automatic') {
        // Start checking for empty folders
        this._startEmptyFolderCheck();
      } else {
        console.log('[ProjectManager] Manual mode enabled, automatic monitoring disabled');
      }
    } else {
      console.log('[ProjectManager] No template settings found, monitoring disabled');
    }

    return Promise.resolve();
  }

  private async _loadTemplateSettings(): Promise<void> {
    try {

      // Get UserCustomActions from web
      const userCustomActions = await this._sp.web.userCustomActions();

      // Cast to our custom interface type
      const customActions = userCustomActions as IAppCustomizerInfo[];

      // Find our extension (ES5 compatible)
      let ourExtension: IAppCustomizerInfo | null = null;
      for (let i = 0; i < customActions.length; i++) {
        if (customActions[i].ClientSideComponentId === EXTENSION_ID) {
          ourExtension = customActions[i];
          break;
        }
      }

      if (ourExtension && ourExtension.ClientSideComponentProperties) {
        try {
          const props = JSON.parse(ourExtension.ClientSideComponentProperties);
          if (props.templateSettings) {
            this._templateSettings = props.templateSettings;
            console.log('[ProjectManager] Loaded template settings:', this._templateSettings);
          }
        } catch (e) {
          console.error('[ProjectManager] Error parsing settings:', e);
        }
      }
    } catch (error) {
      console.error('[ProjectManager] Error loading template settings:', error);
    }
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const templateSettingsCommand = this.tryGetCommand('TEMPLATE_SETTINGS');
    if (templateSettingsCommand) {
      // Show the Template Settings button only in ProjectDocs library or other document libraries
      const pageContext = this.context.pageContext;
      const isInDocumentLibrary = pageContext && pageContext.list && pageContext.list.serverRelativeUrl;

      // You can customize this logic to show the button in specific libraries
      // For now, show it in all document libraries
      templateSettingsCommand.visible = !!isInDocumentLibrary;

    }
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'TEMPLATE_SETTINGS':
        await this._showTemplateSettings();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async _showTemplateSettings(): Promise<void> {
    const pageContext = this.context.pageContext;
    const currentSiteUrl = pageContext.web.absoluteUrl;

    // Get the library internal name from the server relative URL
    let currentLibrary = 'Documents';  // Fallback default
    if (pageContext.list && pageContext.list.serverRelativeUrl) {
      const urlParts = pageContext.list.serverRelativeUrl.split('/');
      currentLibrary = urlParts[urlParts.length - 1] || 'Documents';
    }

    // Get current folder from URL
    const currentUrl = window.location.href;
    let currentFolder = '';

    if (currentUrl.indexOf('RootFolder=') !== -1) {
      const match = currentUrl.match(/RootFolder=([^&]*)/);
      if (match) {
        currentFolder = decodeURIComponent(match[1]);
        // Remove the site and library parts to get just the folder path
        const libraryPath = pageContext.list ? pageContext.list.serverRelativeUrl : '';
        if (currentFolder.indexOf(libraryPath) === 0) {
          currentFolder = currentFolder.substring(libraryPath.length);
          if (currentFolder.indexOf('/') === 0) {
            currentFolder = currentFolder.substring(1);
          }
        }
      }
    } else if (currentUrl.indexOf('/Forms/AllItems.aspx?id=') !== -1) {
      const match = currentUrl.match(/id=([^&]*)/);
      if (match) {
        const fullPath = decodeURIComponent(match[1]);
        const libraryPath = pageContext.list ? pageContext.list.serverRelativeUrl : '';
        if (fullPath.indexOf(libraryPath) === 0) {
          currentFolder = fullPath.substring(libraryPath.length);
          if (currentFolder.indexOf('/') === 0) {
            currentFolder = currentFolder.substring(1);
          }
        }
      }
    }

    console.log('[ProjectManager] Opening Template Settings dialog');
    console.log('[ProjectManager] Current Site:', currentSiteUrl);
    console.log('[ProjectManager] Current Library:', currentLibrary);
    console.log('[ProjectManager] Current Folder:', currentFolder);

    const dialog = new TemplateSettingsDialog(
      this._sp,
      currentSiteUrl,
      currentLibrary,
      currentFolder,
      EXTENSION_ID,
      async (settings: ITemplateSettings) => {
        console.log('[ProjectManager] Template settings saved:', settings);
        // Update local properties
        this.properties.templateSettings = settings;
        this.properties.isConfigured = true;
        // Update the cached settings
        this._templateSettings = settings;

        // Stop existing interval if switching modes
        if (this._checkInterval) {
          window.clearInterval(this._checkInterval);
          this._checkInterval = undefined;
        }

        // Restart monitoring if in automatic mode
        if (settings.copyMode === 'automatic') {
          this._startEmptyFolderCheck();
        }
      },
      // Callback for manual template loading
      () => {
        console.log('[ProjectManager] Manual template load triggered');
        dialog.close();
        this._showTemplateLoader();
      }
    );

    await dialog.show();
  }

  private _getLibraryInternalNameFromUrl(): string | null {
    const pageContext = this.context.pageContext;
    if (!pageContext.list) {
      return null;
    }

    const libraryUrl = pageContext.list.serverRelativeUrl;
    const librarySegments = libraryUrl.split('/');
    return librarySegments[librarySegments.length - 1];
  }

  private async _showTemplateLoader(): Promise<void> {
    // Check if we have template settings
    if (!this._templateSettings) {
      console.error('[ProjectManager] No template settings configured');
      return;
    }

    // Get current location info
    const pageContext = this.context.pageContext;
    const libraryInternalName = this._getLibraryInternalNameFromUrl();

    // Check if we're in the configured target library
    if (libraryInternalName !== this._templateSettings.targetLibrary) {
      console.error('[ProjectManager] Not in the configured target library');
      alert(`Template loading is only available in the ${this._templateSettings.targetLibrary} library.`);
      return;
    }

    // Get folder URL from the page URL
    let folderUrl = '';
    const currentUrl = window.location.href;

    if (currentUrl.indexOf('/Forms/AllItems.aspx') !== -1 && currentUrl.indexOf('RootFolder=') !== -1) {
      const match = currentUrl.match(/RootFolder=([^&]*)/);
      if (match) {
        folderUrl = decodeURIComponent(match[1]);
      }
    } else if (currentUrl.indexOf('/Forms/AllItems.aspx?id=') !== -1) {
      const match = currentUrl.match(/id=([^&]*)/);
      if (match) {
        folderUrl = decodeURIComponent(match[1]);
      }
    }

    if (!folderUrl && pageContext.list) {
      folderUrl = pageContext.list.serverRelativeUrl;
    }

    if (!folderUrl) {
      console.error('[ProjectManager] Could not determine current folder');
      alert('Could not determine the current folder location.');
      return;
    }

    // Check folder depth
    const libraryPath = pageContext.list?.serverRelativeUrl || '';
    const relativePath = folderUrl.substring(libraryPath.length);
    const relativeSegments = relativePath.split('/').filter(s => s.length > 0);
    const requiredDepth = this._templateSettings.projectFolderLevels || 2;

    if (relativeSegments.length !== requiredDepth) {
      const depthMessage = requiredDepth === 1
        ? 'Please navigate to a customer folder (1 level deep).'
        : 'Please navigate to a project folder (2 levels deep - Location > Project).';
      alert(`Template loading is only available at the configured folder level. ${depthMessage}`);
      return;
    }

    // Extract folder info based on depth
    let programName = '';
    let locationName = '';
    let projectName = '';

    if (requiredDepth === 1) {
      projectName = relativeSegments[0];
      locationName = projectName;
      programName = this._templateSettings?.program || '';
    } else if (requiredDepth === 2) {
      locationName = relativeSegments[0];
      projectName = relativeSegments[1];
      programName = this._templateSettings?.program || '';
    }

    console.log('[ProjectManager] Manual template load:');
    console.log('  - Folder URL:', folderUrl);
    console.log('  - Program:', programName);
    console.log('  - Location:', locationName);
    console.log('  - Project:', projectName);

    // Show the template loader dialog (manual mode - don't auto-start)
    const dialog = new TemplateLoaderDialog(
      folderUrl,
      programName,
      locationName,
      projectName,
      async (destFolder: string, onProgress?: (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => void) => {
        try {
          await this._loadTemplate(destFolder, programName, onProgress);
        } catch (error) {
          console.error('[ProjectManager] Error loading template:', error);
          throw error;
        }
      },
      () => {
        console.log('[ProjectManager] Template loading cancelled');
      },
      false  // Don't auto-start in manual mode
    );

    await dialog.show();
  }

  private _startEmptyFolderCheck(): void {
    // Only start automatic checking if in automatic mode
    if (this._templateSettings?.copyMode === 'manual') {
      console.log('[ProjectManager] Manual mode enabled, automatic checking disabled');
      return;
    }

    // Use polling interval from settings, default to 8 seconds
    const intervalSeconds = this._templateSettings?.pollingInterval || 8;
    const intervalMs = intervalSeconds * 1000;

    console.log(`[ProjectManager] Starting automatic folder check every ${intervalSeconds} seconds`);
    this._checkInterval = window.setInterval(() => {
      this._checkCurrentFolder();
    }, intervalMs);

    // Also check immediately on navigation
    this._checkCurrentFolder();
  }

  private async _checkCurrentFolder(): Promise<void> {
    try {
      // Skip if no settings are configured
      if (!this._templateSettings) {
        console.log('[ProjectManager] No template settings, skipping folder check');
        return;
      }

      // Get current location info
      const pageContext = this.context.pageContext;
      if (!pageContext || !pageContext.list) {
        return;
      }

      // Get the library internal name from server relative URL
      const libraryUrl = pageContext.list.serverRelativeUrl;
      const librarySegments = libraryUrl.split('/');
      const libraryInternalName = librarySegments[librarySegments.length - 1];

      // Check if we're in the configured target library
      if (libraryInternalName !== this._templateSettings.targetLibrary) {
        return;
      }

      // Get folder URL from the page URL
      let folderUrl = '';
      const url = window.location.href;

      // Extract folder path from SharePoint URL
      // Pattern 1: /Forms/AllItems.aspx?id=/sites/...
      const idMatch = url.match(/[?&]id=([^&]+)/);
      if (idMatch && idMatch[1]) {
        folderUrl = decodeURIComponent(idMatch[1]);
      }
      // Pattern 2: /Documents/Forms/AllItems.aspx?RootFolder=/sites/...
      else {
        const rootFolderMatch = url.match(/[?&]RootFolder=([^&]+)/);
        if (rootFolderMatch && rootFolderMatch[1]) {
          folderUrl = decodeURIComponent(rootFolderMatch[1]);
        } else {
          // Try to get from pageContext if available
          if (pageContext.site && pageContext.web && pageContext.list) {
            // If we're in a folder view, extract from URL path
            if (url.indexOf('/Forms/') > -1) {
              return; // Can't determine folder from this view
            }
          } else {
            return;
          }
        }
      }

      if (!folderUrl) {
        console.log('[ProjectManager] No folder URL found, skipping check');
        return;
      }


      // Parse the folder path to check if we're at the right level based on settings
      const pathSegments = folderUrl.split('/').filter((s: string) => s);
      console.log('[ProjectManager] Path segments:', pathSegments);

      // Find the index of the target library in the path
      let libraryIndex = -1;
      for (let i = 0; i < pathSegments.length; i++) {
        if (pathSegments[i] === libraryInternalName) {
          libraryIndex = i;
          break;
        }
      }

      if (libraryIndex === -1) {
        console.log('[ProjectManager] Could not find target library in path');
        return;
      }

      // Get segments after the library
      const relativeSegments = pathSegments.slice(libraryIndex + 1);

      // Determine required depth based on projectFolderLevels setting
      const requiredDepth = this._templateSettings.projectFolderLevels || 2; // Default to 2 if not set

      // Check if we're at the correct depth
      if (relativeSegments.length !== requiredDepth) {
        return;
      }

      // Extract folder information based on the structure
      let programName = '';
      let locationName = '';
      let projectName = '';

      if (requiredDepth === 1) {
        // Customer Folder structure: /Library/[CustomerFolder]
        programName = relativeSegments[0];
      } else if (requiredDepth === 2) {
        // Location > Project structure: /Library/[Location]/[Project]
        locationName = relativeSegments[0];
        projectName = relativeSegments[1];

        // For 2-level structure, use the Program from settings if available
        // This should be set in the Template Settings dialog
        if (this._templateSettings?.program) {
          programName = this._templateSettings.program;
        } else {
          // Fallback: try to extract from location name or leave empty
          programName = ''; // Don't use location as program!
          console.log('[ProjectManager] Warning: No Program set in settings, templates will only include General/empty items');
        }
      }

      // Check if we're already loading a template
      if (this._isLoadingTemplate) {
        console.log('[ProjectManager] Template loading already in progress, skipping check');
        return;
      }

      // Skip if we're already checking this folder and it was checked recently
      if (this._lastCheckedFolder === folderUrl) {
        // Check cache to see if we've already determined this folder's state
        const cached = this._folderCheckCache[folderUrl];
        if (cached) {
          const cacheAge = Date.now() - cached.checkedAt;
          // Cache is valid for 30 seconds to avoid repeated checks
          if (cacheAge < 30000) {
            console.log('[ProjectManager] Using cached result for folder, skipping API call');
            return;
          }
        }
      } else {
        // New folder, update tracking
        this._lastCheckedFolder = folderUrl;
        console.log('[ProjectManager] Now monitoring new folder:', folderUrl);
      }

      // Check if dialog has already been shown for this folder
      const dialogKey = `templateDialog_${folderUrl}`;
      const dialogShown = sessionStorage.getItem(dialogKey);
      const templateLoaded = sessionStorage.getItem(`${dialogKey}_loaded`);

      if (dialogShown || templateLoaded) {
        console.log('[ProjectManager] Dialog already shown or template already loaded for this folder');
        // Cache this result to avoid future checks
        this._folderCheckCache[folderUrl] = { isEmpty: false, checkedAt: Date.now() };
        return;
      }

      // Check if folder is empty (this makes the API call)
      const isEmpty = await this._isFolderEmpty(folderUrl);
      console.log('[ProjectManager] Folder empty check result:', isEmpty);

      // Cache the result
      this._folderCheckCache[folderUrl] = { isEmpty, checkedAt: Date.now() };

      if (isEmpty) {
        console.log('[ProjectManager] Showing template dialog for empty folder');
        // Show dialog to load template
        await this._showTemplateDialog(folderUrl, programName, locationName, projectName);
      } else {
        // Mark that this folder already has content
        sessionStorage.setItem(`${dialogKey}_loaded`, 'true');
        console.log('[ProjectManager] Folder not empty, marking as already loaded');
      }
    } catch (error) {
      console.error('Error checking folder:', error);
    }
  }

  private async _isFolderEmpty(folderUrl: string): Promise<boolean> {
    try {
      // Use Lists API with CAML query for better performance
      // Build the folder path relative to the library
      const libraryUrl = this.context.pageContext.list?.serverRelativeUrl;
      let folderPath = folderUrl;

      if (libraryUrl && folderUrl.indexOf(libraryUrl) === 0) {
        // Remove the library URL to get the relative path
        folderPath = folderUrl.substring(libraryUrl.length);
        if (folderPath.indexOf('/') === 0) {
          folderPath = folderPath.substring(1);
        }
      }

      console.log(`[ProjectManager] Checking folder using Lists API: ${folderPath || '/'}`);

      // Use GetFolderByServerRelativeUrl to avoid list view threshold issues
      const folderEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')?$expand=Folders,Files`;

      const response = await this.context.spHttpClient.get(folderEndpoint, SPHttpClient.configurations.v1);

      if (!response.ok) {
        console.log('[ProjectManager] Error checking folder contents');
        return false; // Treat as not empty if we can't check
      }

      const data = await response.json();

      // Check if folder has any files or subfolders (excluding Forms folder)
      const hasFiles = data.Files && data.Files.length > 0;
      const folders = data.Folders ? data.Folders.filter((f: any) => f.Name !== 'Forms') : [];
      const hasSubfolders = folders.length > 0;

      const isEmpty = !hasFiles && !hasSubfolders;
      console.log(`[ProjectManager] Folder "${folderUrl}" - Files: ${data.Files?.length || 0}, Folders: ${folders.length}, Empty: ${isEmpty}`);

      return isEmpty;
    } catch (error) {
      console.error('[ProjectManager] Error checking if folder is empty:', error);
      return false; // Treat as not empty if we can't check
    }
  }

  private async _showTemplateDialog(folderUrl: string, programName: string, locationName: string, projectName: string): Promise<void> {
    // Double-check if dialog has already been shown for this folder
    const dialogKey = `templateDialog_${folderUrl}`;
    if (sessionStorage.getItem(dialogKey)) {
      console.log('[ProjectManager] Dialog already shown for this folder, skipping');
      return;
    }

    // Mark that we've shown the dialog for this folder
    sessionStorage.setItem(dialogKey, 'shown');
    this._isLoadingTemplate = true;

    // Show the template loader dialog (automatic mode - auto-start)
    const dialog = new TemplateLoaderDialog(
      folderUrl,
      programName,
      locationName,
      projectName,
      async (destFolder: string, onProgress?: (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => void) => {
        try {
          await this._loadTemplate(destFolder, programName, onProgress);
          // Mark that template was successfully loaded
          sessionStorage.setItem(`${dialogKey}_loaded`, 'true');
        } finally {
          this._isLoadingTemplate = false;
        }
      },
      () => {
        // On cancel, allow the dialog to be shown again
        sessionStorage.removeItem(dialogKey);
        this._isLoadingTemplate = false;
      },
      true  // Auto-start in automatic mode
    );

    try {
      await dialog.show();
    } catch (error) {
      console.error('[ProjectManager] Error showing dialog:', error);
      this._isLoadingTemplate = false;
      sessionStorage.removeItem(dialogKey);
    }
  }

  private async _loadTemplate(destinationFolderUrl: string, currentProgramName: string, onProgress?: (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => void): Promise<void> {
    try {
      console.log('[ProjectManager] Loading template for program:', currentProgramName);
      console.log('[ProjectManager] Destination folder:', destinationFolderUrl);

      // Get the target library information ONCE for all operations
      const libraryInternalName = this._templateSettings?.targetLibrary || 'Documents';
      let targetListTitle = '';
      let targetListItemType = '';

      try {
        const listEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=RootFolder/Name eq '${libraryInternalName}'&$select=Title,ListItemEntityTypeFullName`;
        const listResponse = await this.context.spHttpClient.get(listEndpoint, SPHttpClient.configurations.v1);

        if (listResponse.ok) {
          const listData = await listResponse.json();
          if (listData.value && listData.value.length > 0) {
            targetListTitle = listData.value[0].Title;
            targetListItemType = listData.value[0].ListItemEntityTypeFullName;
            console.log(`[ProjectManager] Target library: ${targetListTitle} (${libraryInternalName}), Type: ${targetListItemType}`);
          }
        }
      } catch (error) {
        console.error('[ProjectManager] Error getting target library information:', error);
      }

      // Get the folder's metadata fields (Program, ProjectStatus, ProjectNumber)
      let folderMetadata: { Program?: string; ProjectStatus?: string; ProjectNumber?: string } = {};
      try {
        console.log('[ProjectManager] Getting folder metadata from:', destinationFolderUrl);

        // Get the folder's list item to access its metadata
        const folderEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(destinationFolderUrl)}')/ListItemAllFields?$select=Program,ProjectStatus,ProjectNumber`;
        const folderResponse = await this.context.spHttpClient.get(folderEndpoint, SPHttpClient.configurations.v1);

        if (folderResponse.ok) {
          const folderData = await folderResponse.json();

          // Extract metadata, but don't default to empty string - keep as undefined if not present
          // Also check for case variations in field names
          folderMetadata = {
            Program: folderData.Program || folderData.program || undefined,
            ProjectStatus: folderData.ProjectStatus || folderData.projectStatus || folderData.Project_x0020_Status || undefined,
            ProjectNumber: folderData.ProjectNumber || folderData.projectNumber || folderData.Project_x0020_Number || undefined
          };

          console.log('[ProjectManager] Raw folder data from API:', folderData);
          console.log('[ProjectManager] Folder metadata extracted:', folderMetadata);

          // Only use metadata if at least one field has a value
          const hasMetadata = folderMetadata.Program || folderMetadata.ProjectStatus || folderMetadata.ProjectNumber;
          if (hasMetadata) {
            console.log('[ProjectManager] ✅ Will apply metadata to copied items');
          } else {
            console.log('[ProjectManager] ⚠️ No metadata values found in destination folder');
          }
        } else {
          console.log('[ProjectManager] Could not retrieve folder metadata, HTTP status:', folderResponse.status);
        }
      } catch (error) {
        console.log('[ProjectManager] Error getting folder metadata:', error);
        // Continue without metadata - not a critical error
      }

      // Use saved settings from _templateSettings (loaded at initialization)
      let sourceUrl = '/sites/ProjectsRepository/Templates'; // Default fallback

      if (this._templateSettings) {
        const settings = this._templateSettings;
        console.log('[ProjectManager] Using saved template settings:', settings);

        // Build source URL from settings
        if (settings.sourceSite && settings.sourceLibrary) {
          let siteUrl = settings.sourceSite;
          // Remove trailing slash if present
          if (siteUrl.charAt(siteUrl.length - 1) === '/') {
            siteUrl = siteUrl.slice(0, -1);
          }
          // Extract site path from full URL if needed
          if (siteUrl.indexOf('.sharepoint.com') !== -1) {
            const match = siteUrl.match(/\.sharepoint\.com(.+)/);
            if (match) {
              siteUrl = match[1];
            }
          }

          sourceUrl = `${siteUrl}/${settings.sourceLibrary}`;
          if (settings.sourceFolder) {
            sourceUrl = `${sourceUrl}/${settings.sourceFolder}`;
          }

          console.log('[ProjectManager] Constructed source URL from settings:', sourceUrl);

          // Also use the program from settings if available for filtering
          if (settings.program && settings.program !== 'All') {
            console.log('[ProjectManager] Will filter templates by program:', settings.program);
            // Override the currentProgramName with the selected program from settings
            currentProgramName = settings.program;
          }
        }
      } else {
        console.log('[ProjectManager] No saved settings, using default source:', sourceUrl);
      }

      // Get all items from Templates library with metadata
      const items = await this._getAllItemsFromFolderWithMetadata(sourceUrl);

      console.log('[ProjectManager] Retrieved items from Templates library:');
      console.log('[ProjectManager] Total folders found:', items.folders.length);
      console.log('[ProjectManager] Folder list:', items.folders.map(f => `${f.Name} (${f.ServerRelativeUrl})`));

      if (items.files.length === 0 && items.folders.length === 0) {
        throw new Error('No template files found in the Templates library');
      }

      // Filter items based on Program field
      console.log(`[ProjectManager] ========== FILTERING BY PROGRAM ==========`);
      console.log(`[ProjectManager] Current Program Name: "${currentProgramName}" (type: ${typeof currentProgramName}, length: ${currentProgramName?.length})`);
      console.log(`[ProjectManager] Template Settings:`, this._templateSettings);

      // DEBUG: Show first few items BEFORE filtering
      console.log(`[ProjectManager] BEFORE FILTER - Sample folders:`, items.folders.slice(0, 5).map(f => ({
        Name: f.Name,
        Program: f.Program,
        ProgramType: typeof f.Program,
        ProgramLength: f.Program?.length
      })));

      // First pass: identify folders that should be excluded based on their Program field
      const excludedByProgram = new Set<string>();

      console.log(`[ProjectManager] ========== FIRST PASS: IDENTIFYING EXCLUDED FOLDERS ==========`);
      console.log(`[ProjectManager] Current Program: "${currentProgramName}"`);

      items.folders.forEach(folder => {
        const programValue = folder.Program;

        // If folder has a specific program that doesn't match current program, exclude it
        if (programValue &&
          programValue !== '' &&
          programValue !== null &&
          programValue !== undefined &&
          typeof programValue === 'string' &&
          programValue.toLowerCase() !== 'general' &&
          programValue !== currentProgramName) {

          excludedByProgram.add(folder.ServerRelativeUrl);
          console.log(`[ProjectManager] ❌ Excluding folder "${folder.Name}" at path "${folder.ServerRelativeUrl}" (Program: "${programValue}" doesn't match current: "${currentProgramName}")`);
        } else {
          console.log(`[ProjectManager] ✓ Not excluding "${folder.Name}" (Program: "${programValue}")`);
        }
      });

      console.log(`[ProjectManager] Excluded folders count: ${excludedByProgram.size}`);
      const excludedArray: string[] = [];
      excludedByProgram.forEach(path => excludedArray.push(path));
      console.log(`[ProjectManager] Excluded paths:`, excludedArray);

      // Second pass: filter folders, excluding those under excluded parents
      console.log(`[ProjectManager] ========== SECOND PASS: FILTERING FOLDERS ==========`);

      const filteredFolders = items.folders.filter(folder => {
        const programValue = folder.Program;

        // DETAILED DEBUG
        console.log(`[ProjectManager] ====== EVALUATING FOLDER: "${folder.Name}" ======`);
        console.log(`[ProjectManager]   - Full Path: "${folder.ServerRelativeUrl}"`);
        console.log(`[ProjectManager]   - Program value: "${programValue}"`);
        console.log(`[ProjectManager]   - Current program: "${currentProgramName}"`);

        // Check if this folder is under any excluded parent
        let isUnderExcludedParent = false;
        excludedByProgram.forEach(excludedPath => {
          // Check if this folder is a child of an excluded folder
          // folder.ServerRelativeUrl should start with excludedPath + '/'
          // Example: "/sites/repo/Docs/Amazon/Whatever" starts with "/sites/repo/Docs/Amazon/"
          if (folder.ServerRelativeUrl.indexOf(excludedPath + '/') === 0) {
            isUnderExcludedParent = true;
            console.log(`[ProjectManager]   - Under excluded parent: "${excludedPath}"`);
          }
          // Also check if this IS the excluded folder itself
          if (folder.ServerRelativeUrl === excludedPath) {
            isUnderExcludedParent = true;
            console.log(`[ProjectManager]   - This IS an excluded folder: "${excludedPath}"`);
          }
        });

        if (isUnderExcludedParent) {
          console.log(`[ProjectManager] → DECISION: ❌❌❌ EXCLUDE (parent excluded) - "${folder.Name}"`);
          return false;
        }

        // Check folder's own Program field
        const folderProgramMatches = !programValue ||
          programValue === '' ||
          programValue === null ||
          programValue === undefined ||
          (typeof programValue === 'string' && programValue.toLowerCase() === 'general') ||
          programValue === currentProgramName;

        // Special case: If folder name is a known program name but Program field is empty,
        // treat it as if it has that Program value
        const knownPrograms = ['Amazon', 'FedEx', 'UPS', 'DHL'];
        let impliedProgram = null;
        for (let i = 0; i < knownPrograms.length; i++) {
          if (folder.Name === knownPrograms[i] && !programValue) {
            impliedProgram = knownPrograms[i];
            console.log(`[ProjectManager]   - WARNING: Folder name "${folder.Name}" matches a program but Program field is empty`);
            console.log(`[ProjectManager]   - Treating as Program: "${impliedProgram}"`);
            break;
          }
        }

        // If folder has implied program that doesn't match current, exclude it
        const shouldCopy = !impliedProgram ? folderProgramMatches : (impliedProgram === currentProgramName);

        console.log(`[ProjectManager] → DECISION: ${shouldCopy ? '✅✅✅ INCLUDE' : '❌❌❌ EXCLUDE'} - "${folder.Name}"`);

        // ALERT if we're including a FedEx item when current program is not FedEx
        if (shouldCopy && programValue === 'FedEx' && currentProgramName !== 'FedEx') {
          console.error(`[ProjectManager] 🚨🚨🚨 ERROR: INCLUDING FedEx item "${folder.Name}" when current program is "${currentProgramName}"!`);
        }

        return shouldCopy;
      });

      // Create a set of allowed folder paths (folders that passed the filter)
      const allowedFolderPaths = new Set<string>();
      filteredFolders.forEach(folder => {
        allowedFolderPaths.add(folder.ServerRelativeUrl);
      });

      // Also need to check parent folders - if parent is excluded, exclude all children
      const excludedFolderPaths = new Set<string>();

      items.folders.forEach(folder => {
        // Check if folder is in filteredFolders using indexOf
        let isInFilteredFolders = false;
        for (let i = 0; i < filteredFolders.length; i++) {
          if (filteredFolders[i] === folder) {
            isInFilteredFolders = true;
            break;
          }
        }
        if (!isInFilteredFolders) {
          excludedFolderPaths.add(folder.ServerRelativeUrl);
        }
      });

      // Convert Sets to arrays for logging (ES5 compatible)
      const allowedFolderPathsArray: string[] = [];
      allowedFolderPaths.forEach(path => allowedFolderPathsArray.push(path));
      const excludedFolderPathsArray: string[] = [];
      excludedFolderPaths.forEach(path => excludedFolderPathsArray.push(path));

      console.log(`[ProjectManager] Allowed folder paths (${allowedFolderPaths.size}):`, allowedFolderPathsArray.slice(0, 5));
      console.log(`[ProjectManager] Excluded folder paths (${excludedFolderPaths.size}):`, excludedFolderPathsArray.slice(0, 5));

      const filteredFiles = items.files.filter(file => {
        const programValue = file.Program;

        // First check if the file's own Program field matches
        const fileProgramMatches = !programValue ||
          programValue === '' ||
          programValue === null ||
          programValue === undefined ||
          (typeof programValue === 'string' && programValue.toLowerCase() === 'general') ||
          programValue === currentProgramName;

        // IMPORTANT: Also check if the file's parent folder is in the allowed list
        // Extract parent folder path from file path
        const parentFolderPath = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf('/'));

        // Check if the file is under an excluded folder (at any level)
        let isUnderExcludedFolder = false;
        excludedFolderPaths.forEach(excludedPath => {
          // Check if the file is inside this excluded folder
          // This handles both direct children and nested descendants
          if (file.ServerRelativeUrl.indexOf(excludedPath + '/') > -1) {
            isUnderExcludedFolder = true;
            console.log(`[ProjectManager] File "${file.Name}" is under excluded folder: ${excludedPath}`);
          }
        });

        // Also check if any parent folder in the hierarchy was excluded
        // This ensures we catch all cases where a parent folder doesn't match the program
        let currentPath = parentFolderPath;
        while (currentPath && currentPath.length > sourceUrl.length && !isUnderExcludedFolder) {
          if (excludedFolderPaths.has(currentPath)) {
            isUnderExcludedFolder = true;
            console.log(`[ProjectManager] File "${file.Name}" excluded - parent folder "${currentPath}" was filtered out`);
            break;
          }
          // Move up one level
          const lastSlash = currentPath.lastIndexOf('/');
          if (lastSlash > 0) {
            currentPath = currentPath.substring(0, lastSlash);
          } else {
            break;
          }
        }

        // File must match program AND NOT be under an excluded folder
        const shouldCopy = fileProgramMatches && !isUnderExcludedFolder;

        // Only log first few files to avoid spam
        if (items.files.indexOf(file) < 10) {
          console.log(`[ProjectManager] File "${file.Name}"`);
          console.log(`[ProjectManager]   - File Program: "${programValue}" → Match: ${fileProgramMatches}`);
          console.log(`[ProjectManager]   - Parent folder: "${parentFolderPath}"`);
          console.log(`[ProjectManager]   - Under excluded folder: ${isUnderExcludedFolder}`);
          console.log(`[ProjectManager]   → ${shouldCopy ? '✅ INCLUDE' : '❌ EXCLUDE'}`);
        }

        // Alert if we're excluding a file because its parent folder was filtered out
        if (fileProgramMatches && isUnderExcludedFolder) {
          console.log(`[ProjectManager] ⚠️ File "${file.Name}" excluded because parent folder was filtered out`);
        }

        return shouldCopy;
      });

      console.log(`[ProjectManager] ========== FILTERING RESULTS ==========`);
      console.log(`[ProjectManager] Total folders: ${items.folders.length} → Filtered: ${filteredFolders.length} (removed ${items.folders.length - filteredFolders.length})`);
      console.log(`[ProjectManager] Total files: ${items.files.length} → Filtered: ${filteredFiles.length} (removed ${items.files.length - filteredFiles.length})`);
      console.log('[ProjectManager] Filtered folder names:', filteredFolders.map(f => `${f.Name} (Program: ${f.Program})`).slice(0, 10));

      // Check if we're incorrectly including items
      const wrongPrograms = filteredFolders.filter(f =>
        f.Program &&
        f.Program !== '' &&
        f.Program !== currentProgramName &&
        f.Program.toLowerCase() !== 'general'
      );
      if (wrongPrograms.length > 0) {
        console.error(`[ProjectManager] ⚠️ WARNING: Found ${wrongPrograms.length} folders with wrong program that shouldn't be included:`,
          wrongPrograms.map(f => `${f.Name} (Program: ${f.Program})`));
      }

      // Remove duplicate folders by tracking relative paths
      const seenPaths: { [key: string]: boolean } = {};
      const uniqueFolders: IFolderItem[] = [];

      for (let i = 0; i < filteredFolders.length; i++) {
        const folder = filteredFolders[i];
        const relativePath = folder.ServerRelativeUrl.replace(sourceUrl, '');
        console.log(`[ProjectManager] Processing folder "${folder.Name}" with relative path: "${relativePath}"`);

        if (!seenPaths[relativePath]) {
          seenPaths[relativePath] = true;
          uniqueFolders.push(folder);
          console.log(`[ProjectManager] Added unique folder: ${folder.Name}`);
        } else {
          console.log(`[ProjectManager] Skipping duplicate folder: ${folder.Name} (path: ${relativePath})`);
        }
      }

      // Sort folders by depth to ensure parent folders are created before children
      uniqueFolders.sort((a: IFolderItem, b: IFolderItem) => {
        const aDepth = a.ServerRelativeUrl.split('/').length;
        const bDepth = b.ServerRelativeUrl.split('/').length;
        return aDepth - bDepth;
      });

      console.log(`[ProjectManager] After deduplication: ${uniqueFolders.length} unique folders and ${filteredFiles.length} files for program "${currentProgramName}"`);

      const totalItems = uniqueFolders.length + filteredFiles.length;

      if (totalItems === 0) {
        throw new Error(`No template files found for program "${currentProgramName}"`);
      }

      let processedItems = 0;

      // Track success/failure
      let foldersCreated = 0;
      let foldersFailed = 0;
      let filesCopied = 0;
      let filesFailed = 0;

      // Copy unique folders first (sorted by depth)
      console.log('[ProjectManager] ========== CREATING FOLDER STRUCTURE ==========');
      console.log(`[ProjectManager] Total folders to create: ${uniqueFolders.length}`);

      for (const folder of uniqueFolders) {
        const relativePath = folder.ServerRelativeUrl.replace(sourceUrl, '');
        const destFolderUrl = `${destinationFolderUrl}${relativePath}`;

        // Extract just the folder structure after the library for display
        const displayPath = relativePath.indexOf('/') === 0 ? relativePath.substring(1) : relativePath;

        if (onProgress) {
          onProgress(
            Math.round((processedItems / totalItems) * 100),
            `Creating folder: ${displayPath || folder.Name}`,
            processedItems + 1,
            totalItems
          );
        }

        console.log(`[ProjectManager] [${foldersCreated + foldersFailed + 1}/${uniqueFolders.length}] Processing folder: "${folder.Name}"`);

        const success = await this._createFolderPath(destFolderUrl);
        if (!success) {
          // STOP IMMEDIATELY on failure
          const errorMsg = `Failed to create folder: "${folder.Name}" at path: ${destFolderUrl}`;
          console.error(`[ProjectManager] ❌ CRITICAL ERROR: ${errorMsg}`);
          throw new Error(errorMsg);
        }

        // Update the created folder's metadata IMMEDIATELY after creation
        if (folderMetadata && (folderMetadata.Program || folderMetadata.ProjectStatus || folderMetadata.ProjectNumber)) {
          try {
            console.log(`[ProjectManager] Setting metadata on newly created folder: ${destFolderUrl}`);

            // Get the folder's list item ID using REST API
            const folderEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(destFolderUrl)}')/ListItemAllFields?$select=ID,FileSystemObjectType`;
            const folderResponse = await this.context.spHttpClient.get(folderEndpoint, SPHttpClient.configurations.v1);

            if (folderResponse.ok) {
              const folderData = await folderResponse.json();
              const itemId = folderData.ID || folderData.Id;

              if (itemId) {
                console.log(`[ProjectManager] Got folder list item ID: ${itemId}`);

                // Use the already-fetched library information
                if (targetListTitle && targetListItemType) {
                  // Build the update payload
                  const updatePayload: any = {
                    "__metadata": { "type": targetListItemType }
                  };

                  if (folderMetadata.Program) updatePayload.Program = folderMetadata.Program;
                  if (folderMetadata.ProjectStatus) updatePayload.ProjectStatus = folderMetadata.ProjectStatus;
                  if (folderMetadata.ProjectNumber) updatePayload.ProjectNumber = folderMetadata.ProjectNumber;

                  // Update the folder's metadata using REST API with SPHttpClient
                  const updateEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(targetListTitle)}')/items(${itemId})`;

                  const headers: any = {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-HTTP-Method': 'MERGE',
                    'IF-MATCH': '*',
                    'odata-version': ''
                  };

                  const updateResponse = await this.context.spHttpClient.post(
                    updateEndpoint,
                    SPHttpClient.configurations.v1,
                    {
                      headers: headers,
                      body: JSON.stringify(updatePayload)
                    }
                  );

                  if (updateResponse.ok) {
                    console.log(`[ProjectManager] ✅ Metadata set on folder: ${destFolderUrl}`);
                  } else {
                    const errorText = await updateResponse.text();
                    console.error(`[ProjectManager] Failed to update folder metadata. Status: ${updateResponse.status}`, errorText);
                  }
                } else {
                  console.log('[ProjectManager] Cannot update folder metadata - library information not available');
                }
              }
            }
          } catch (error) {
            console.error(`[ProjectManager] Warning: Could not set metadata on folder ${destFolderUrl}:`, error);
            // Continue - don't fail the whole operation
          }
        }

        foldersCreated++;
        processedItems++;
      }

      console.log(`[ProjectManager] Folder creation complete: ${foldersCreated} folders created successfully`);

      // Now copy files after all folders are created
      console.log('[ProjectManager] ========== COPYING FILES ==========');
      console.log(`[ProjectManager] Total files to copy: ${filteredFiles.length}`);

      // Add delay to ensure folders are fully created in SharePoint
      if (filteredFiles.length > 0) {
        console.log('[ProjectManager] Waiting 2 seconds for folders to be fully available...');
        await new Promise(resolve => setTimeout(resolve, 2000));
      }

      // Process files in batches to avoid throttling
      const BATCH_SIZE = 5; // Process 5 files at a time
      const BATCH_DELAY = 1000; // 1 second delay between batches

      for (let i = 0; i < filteredFiles.length; i += BATCH_SIZE) {
        const batch = filteredFiles.slice(i, Math.min(i + BATCH_SIZE, filteredFiles.length));
        console.log(`[ProjectManager] Processing batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(filteredFiles.length / BATCH_SIZE)} (${batch.length} files)`);

        // Process files in the current batch sequentially to avoid overwhelming SharePoint
        for (const file of batch) {
          const relativePath = file.ServerRelativeUrl.replace(sourceUrl, '');
          const destFileUrl = `${destinationFolderUrl}${relativePath}`;

          // Extract folder path for display
          const displayPath = relativePath.indexOf('/') === 0 ? relativePath.substring(1) : relativePath;

          if (onProgress) {
            onProgress(
              Math.round((processedItems / totalItems) * 100),
              `Copying: ${displayPath}`,
              processedItems + 1,
              totalItems
            );
          }

          // Check if the parent folder was in the allowed list (filteredFolders)
          const parentFolderPath = destFileUrl.substring(0, destFileUrl.lastIndexOf('/'));
          const sourceParentPath = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf('/'));

          // Check if this file's parent folder was in our filtered (allowed) folders
          // Also check if the parent is the source root itself
          let parentWasAllowed = sourceParentPath === sourceUrl;

          // ES5 compatible check for parent folder in filteredFolders
          if (!parentWasAllowed) {
            for (let i = 0; i < filteredFolders.length; i++) {
              if (filteredFolders[i].ServerRelativeUrl === sourceParentPath) {
                parentWasAllowed = true;
                break;
              }
            }
          }

          if (parentWasAllowed) {
            console.log(`[ProjectManager] Ensuring parent folder exists for file: ${parentFolderPath}`);
            const folderCreated = await this._ensureFolderExists(parentFolderPath);
            if (!folderCreated) {
              // STOP IMMEDIATELY on folder creation failure
              const errorMsg = `Failed to ensure parent folder for file "${file.Name}". Parent folder path: ${parentFolderPath}`;
              console.error(`[ProjectManager] ❌ CRITICAL ERROR: ${errorMsg}`);
              throw new Error(errorMsg);
            }
          } else {
            // This shouldn't happen if filtering is working correctly
            console.error(`[ProjectManager] ⚠️ WARNING: File "${file.Name}" passed filter but its parent folder was excluded!`);
            console.error(`[ProjectManager]   - File path: ${file.ServerRelativeUrl}`);
            console.error(`[ProjectManager]   - Parent folder: ${sourceParentPath}`);
            console.error(`[ProjectManager]   - This file will be skipped to prevent creating excluded folders`);

            // Skip this file
            filesFailed++;
            processedItems++;
            continue;
          }

          console.log(`[ProjectManager] [${filesCopied + filesFailed + 1}/${filteredFiles.length}] Processing file: "${file.Name}"`);

          // Pass throttling callback to show status in dialog
          const success = await this._copyFile(
            file.ServerRelativeUrl,
            destFileUrl,
            0,
            (isThrottled: boolean, retryInfo?: string) => {
              if (onProgress) {
                onProgress(
                  Math.round((processedItems / totalItems) * 100),
                  retryInfo || `Copying: ${displayPath}`,
                  processedItems + 1,
                  totalItems,
                  isThrottled
                );
              }
            }
          );

          if (!success) {
            // STOP IMMEDIATELY on file copy failure
            const errorMsg = `Failed to copy file: "${file.Name}" from ${file.ServerRelativeUrl} to ${destFileUrl}`;
            console.error(`[ProjectManager] ❌ CRITICAL ERROR: ${errorMsg}`);
            throw new Error(errorMsg);
          }

          // Update the copied file's metadata with folder values if we have them
          if (folderMetadata && (folderMetadata.Program || folderMetadata.ProjectStatus || folderMetadata.ProjectNumber)) {
            // Small delay to let SharePoint register the file
            await new Promise(resolve => setTimeout(resolve, 100)); // 100ms delay

            console.log(`[ProjectManager] Applying folder metadata to file: ${file.Name}`);
            await this._updateFileMetadata(destFileUrl, folderMetadata, targetListTitle, targetListItemType);
          }

          filesCopied++;
          processedItems++;
        }

        // Add delay between batches to avoid throttling (except for the last batch)
        if (i + BATCH_SIZE < filteredFiles.length) {
          console.log(`[ProjectManager] Waiting ${BATCH_DELAY}ms before next batch to avoid throttling...`);
          await new Promise(resolve => setTimeout(resolve, BATCH_DELAY));
        }
      }

      console.log(`[ProjectManager] File copying complete: ${filesCopied} files copied successfully`);

      // Final summary
      console.log('[ProjectManager] ========== TEMPLATE LOADING SUMMARY ==========');
      console.log(`[ProjectManager] ✅ Successfully created ${foldersCreated} folders`);
      console.log(`[ProjectManager] ✅ Successfully copied ${filesCopied} files`);
      console.log(`[ProjectManager] ✅ All operations completed successfully!`);

      // Log the successful creation to the SharePoint list
      const totalItemsCopied = foldersCreated + filesCopied;
      await this._logProjectCreation(
        destinationFolderUrl,
        currentProgramName,
        'Success',
        totalItemsCopied,
        null
      );

      if (onProgress) {
        onProgress(100, 'Template loaded successfully!', totalItems, totalItems);
      }

    } catch (error) {
      console.error('Error loading template:', error);

      // Log the failed creation to the SharePoint list
      await this._logProjectCreation(
        destinationFolderUrl,
        currentProgramName,
        'Failed',
        0,
        error.message || error.toString()
      );

      throw error;
    }
  }

  private async _logProjectCreation(
    projectPath: string,
    program: string,
    status: 'Success' | 'Failed',
    itemsCopied: number,
    errorDetails: string | null
  ): Promise<void> {
    try {
      console.log('[ProjectManager] Logging project creation to SharePoint list...');

      // Extract project name from path
      const pathSegments = projectPath.split('/').filter(s => s.length > 0);
      const projectName = pathSegments.length > 0 ? pathSegments[pathSegments.length - 1] : 'Unknown Project';

      // Create the log entry
      const logEntry = {
        Title: projectName,
        Program: program || 'Not Specified',
        Status: status,
        ItemsCopied: itemsCopied,
        CreatedDateTime: new Date().toISOString(),
        ErrorDetails: errorDetails || ''
      };

      // Add to the ProjectCreationLogs list
      const list = this._sp.web.lists.getByTitle('ProjectCreationLogs');
      await list.items.add(logEntry);

      console.log(`[ProjectManager] ✅ Project creation logged: ${projectName} - ${status}`);

    } catch (error) {
      // Don't throw - logging failure shouldn't break the main process
      console.error('[ProjectManager] Error logging project creation:', error);
    }
  }

  private async _getAllItemsFromFolderWithMetadata(folderUrl: string, visitedPaths?: { [key: string]: boolean }): Promise<{ files: IFileItem[], folders: IFolderItem[] }> {
    try {
      console.log(`[ProjectManager] Getting ALL items from library using single Lists API call: ${folderUrl}`);

      // Build the source site URL
      let sourceSiteUrl = this.context.pageContext.web.absoluteUrl; // Default to current site

      if (this._templateSettings?.sourceSite) {
        // Use the source site from settings
        sourceSiteUrl = this._templateSettings.sourceSite;
        // Ensure it's a full URL
        if (sourceSiteUrl.indexOf('http') !== 0) {
          sourceSiteUrl = `https://${this.context.pageContext.web.serverRelativeUrl.split('/')[0]}${sourceSiteUrl}`;
        }
      }

      // Get the library using its internal name from settings
      const libraryInternalName = this._templateSettings?.sourceLibrary || 'ProjectDocsTemplates';

      // First get the list by its internal name (RootFolder/Name)
      const listEndpoint = `${sourceSiteUrl}/_api/web/lists?$filter=RootFolder/Name eq '${libraryInternalName}'&$select=Id,Title`;
      const listResponse = await this.context.spHttpClient.get(listEndpoint, SPHttpClient.configurations.v1);

      if (!listResponse.ok) {
        console.error('[ProjectManager] Could not find source list by internal name');
        throw new Error('Source library not found');
      }

      const listData = await listResponse.json();
      if (!listData.value || listData.value.length === 0) {
        console.error('[ProjectManager] Source list not found by internal name');
        throw new Error('Source library not found');
      }

      const listId = listData.value[0].Id;
      const libraryTitle = listData.value[0].Title; // Get the actual display title for logging

      console.log(`[ProjectManager] Found source library: "${libraryTitle}" (ID: ${listId})`);

      // Now use the list ID to get ALL items in the library
      const itemsEndpoint = `${sourceSiteUrl}/_api/web/lists(guid'${listId}')/items` +
        `?$select=ID,FileRef,FileLeafRef,FSObjType,Created,Modified,FileSizeDisplay,UniqueId,Program,FileDirRef` +
        `&$top=5000` + // Get up to 5000 items in one call
        `&$orderby=FSObjType desc,FileDirRef,FileLeafRef`; // Folders first, then files, ordered by path

      console.log(`[ProjectManager] Fetching all template items from: ${itemsEndpoint}`);

      const response = await this.context.spHttpClient.get(itemsEndpoint, SPHttpClient.configurations.v1);

      if (!response.ok) {
        console.error('[ProjectManager] Lists API failed');
        throw new Error('Failed to retrieve template items');
      }

      const data = await response.json();
      console.log(`[ProjectManager] Retrieved ${data.value.length} total items from library`);

      const allFiles: IFileItem[] = [];
      const allFolders: IFolderItem[] = [];
      const seenFolders = new Set<string>();

      // If we have a specific source folder from settings, filter to only items in that folder
      let sourceBasePath = folderUrl;
      if (this._templateSettings?.sourceFolder) {
        // Construct the full path including the source folder
        const libPath = `/sites/${sourceSiteUrl.split('/sites/')[1]}/${libraryTitle}`;
        sourceBasePath = `${libPath}/${this._templateSettings.sourceFolder}`;
      }

      console.log(`[ProjectManager] Filtering items from source path: ${sourceBasePath}`);

      for (const item of data.value) {
        // Only include items that are within our source folder
        if (!item.FileRef.startsWith(sourceBasePath)) {
          continue;
        }

        // FSObjType: 0 = File, 1 = Folder
        if (item.FSObjType === 0) {
          // It's a file
          allFiles.push({
            Name: item.FileLeafRef,
            ServerRelativeUrl: item.FileRef,
            TimeCreated: item.Created,
            TimeLastModified: item.Modified,
            Length: parseInt(item.FileSizeDisplay || '0', 10) || 0,
            UniqueId: item.UniqueId,
            Program: item.Program // Don't default to empty string - keep as is
          });
        } else if (item.FSObjType === 1) {
          // It's a folder - check if we haven't seen it yet and it's not Forms
          if (!seenFolders.has(item.FileRef) && item.FileLeafRef !== 'Forms') {
            seenFolders.add(item.FileRef);
            allFolders.push({
              Name: item.FileLeafRef,
              ServerRelativeUrl: item.FileRef,
              TimeCreated: item.Created,
              TimeLastModified: item.Modified,
              ItemCount: 0, // Not available via this API
              UniqueId: item.UniqueId,
              Program: item.Program // Don't default to empty string - keep as is
            });
          }
        }
      }

      console.log(`[ProjectManager] After filtering - Found ${allFiles.length} files and ${allFolders.length} folders with Program metadata`);
      console.log(`[ProjectManager] Sample items with Program field:`,
        allFiles.slice(0, 3).map(f => ({ name: f.Name, program: f.Program })),
        allFolders.slice(0, 3).map(f => ({ name: f.Name, program: f.Program }))
      );

      return { files: allFiles, folders: allFolders };
    } catch (error) {
      console.error('[ProjectManager] Error getting items via Lists API:', error);
      throw error;
    }
  }

  private async _ensureFolderExists(folderPath: string): Promise<boolean> {
    try {
      console.log(`[ProjectManager] Ensuring folder exists: ${folderPath}`);

      // First check if folder exists
      try {
        const folder = await this._sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
        if (folder.Exists) {
          console.log(`[ProjectManager] ✅ Folder already exists: ${folderPath}`);
          return true;
        }
      } catch {
        console.log(`[ProjectManager] Folder doesn't exist, will create: ${folderPath}`);
      }

      // Split path and create recursively
      const segments = folderPath.split('/').filter(s => s);
      let currentPath = '';

      for (let i = 0; i < segments.length; i++) {
        // Build path with leading slash
        if (i === 0) {
          currentPath = '/' + segments[i];
        } else {
          currentPath = currentPath + '/' + segments[i];
        }

        // Skip root segments (sites/ProjectsRepository/Shared Documents)
        if (i < 3) {
          console.log(`[ProjectManager] Skipping root segment ${i}: ${currentPath}`);
          continue;
        }

        console.log(`[ProjectManager] Checking segment ${i}: ${currentPath}`);

        // Check if this segment exists
        try {
          const segmentFolder = await this._sp.web.getFolderByServerRelativePath(currentPath).select('Exists')();
          if (segmentFolder.Exists) {
            console.log(`[ProjectManager] ✓ Segment exists: ${currentPath}`);
            continue;
          }
        } catch {
          console.log(`[ProjectManager] Segment doesn't exist, will create: ${currentPath}`);
        }

        // Create the folder segment
        const parentPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
        const folderName = segments[i];

        console.log(`[ProjectManager] Creating folder segment "${folderName}" in parent "${parentPath}"`);

        try {
          console.log(`[ProjectManager] Getting parent folder object for: ${parentPath}`);

          // First ensure parent exists
          const parentFolder = await this._sp.web.getFolderByServerRelativePath(parentPath)();
          console.log(`[ProjectManager] Parent folder exists with ServerRelativeUrl: ${parentFolder.ServerRelativeUrl}`);

          // Use correct PnP.js method
          console.log(`[ProjectManager] Calling folders.addUsingPath for: "${folderName}"`);
          const newFolder = await this._sp.web.getFolderByServerRelativePath(parentPath).folders.addUsingPath(folderName);
          const newFolderPath = parentPath + "/" + folderName;

          const folderItem = await this._sp.web.getFolderByServerRelativePath(newFolderPath).listItemAllFields();
          console.log(`⚠️ DEBUG ⚠️ Folder item after creation:`, folderItem);

          console.log(`[ProjectManager] folders.addUsingPath result:`, newFolder);
          console.log(`[ProjectManager] ✅ Created folder segment: ${currentPath}`);
        } catch (createError) {
          console.error(`[ProjectManager] Full error object:`, createError);

          // Extract SharePoint specific error if available
          let spError = '';
          if (createError.data && createError.data.responseBody) {
            try {
              const errorBody = JSON.parse(createError.data.responseBody);
              if (errorBody.error && errorBody.error.message) {
                spError = `\nSharePoint Error: ${errorBody.error.message}`;
              }
            } catch (e) {
              // Ignore parse errors
            }
          }

          const errorDetails = `Failed to create folder segment: ${currentPath}\n` +
            `Parent path: ${parentPath}\n` +
            `Folder name: ${folderName}\n` +
            `Error: ${createError.message || createError}${spError}\n` +
            `Stack: ${createError.stack || 'No stack trace available'}`;
          console.error(`[ProjectManager] ❌ ${errorDetails}`);
          throw new Error(errorDetails);
        }
      }

      // Final verification
      try {
        const finalCheck = await this._sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
        if (finalCheck.Exists) {
          console.log(`[ProjectManager] ✅ Folder path verified: ${folderPath}`);
          return true;
        } else {
          console.error(`[ProjectManager] ❌ Final verification failed: ${folderPath}`);
          return false;
        }
      } catch {
        console.error(`[ProjectManager] ❌ Could not verify folder: ${folderPath}`);
        return false;
      }
    } catch (error) {
      console.error(`[ProjectManager] ❌ Error ensuring folder exists: ${folderPath}`, error);
      return false;
    }
  }

  private async _createFolderPath(folderPath: string): Promise<boolean> {
    try {
      console.log(`[ProjectManager] Attempting to create folder path: ${folderPath}`);

      // Check if folder already exists
      try {
        const folder = await this._sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
        if (folder.Exists) {
          console.log(`[ProjectManager] ✓ Folder already exists: ${folderPath}`);
          return true;
        }
      } catch {
        console.log(`[ProjectManager] Folder doesn't exist, will create: ${folderPath}`);
      }

      // Split the path and create folders recursively
      const segments = folderPath.split('/').filter(s => s);
      let currentPath = '';

      console.log(`[ProjectManager] Path segments to process:`, segments);

      for (let i = 0; i < segments.length; i++) {
        // Build path with leading slash
        if (i === 0) {
          currentPath = '/' + segments[i];
        } else {
          currentPath = currentPath + '/' + segments[i];
        }

        // Skip root segments that are part of the site structure
        if (i < 3) {
          console.log(`[ProjectManager] Skipping root segment ${i}: ${currentPath}`);
          continue; // Skip "sites/ProjectsRepository/Shared Documents"
        }

        console.log(`[ProjectManager] Processing segment ${i}: ${currentPath}`);

        try {
          const folder = await this._sp.web.getFolderByServerRelativePath(currentPath).select('Exists')();
          if (folder.Exists) {
            console.log(`[ProjectManager] ✓ Segment already exists: ${currentPath}`);
            continue; // Move to next segment
          }
        } catch (checkError) {
          console.log(`[ProjectManager] Segment doesn't exist, will create: ${currentPath}`);
        }

        // Folder doesn't exist, create it
        const parentPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
        const folderName = segments[i];

        console.log(`[ProjectManager] Creating folder "${folderName}" in parent "${parentPath}"`);

        try {
          console.log(`[ProjectManager] Getting parent folder object for: ${parentPath}`);

          // First ensure parent folder exists and get it
          const parentFolder = await this._sp.web.getFolderByServerRelativePath(parentPath)();
          console.log(`[ProjectManager] Parent folder retrieved successfully`);
          console.log(`[ProjectManager] Parent folder ServerRelativeUrl: ${parentFolder.ServerRelativeUrl}`);

          // Use the correct PnP.js method: folders.addUsingPath
          console.log(`[ProjectManager] Calling folders.addUsingPath with folderName: "${folderName}"`);

          const newFolder = await this._sp.web.getFolderByServerRelativePath(parentPath).folders.addUsingPath(folderName);

          const newFolderPath = parentPath + "/" + folderName;
          const folderItem: any = await this._sp.web.getFolderByServerRelativePath(newFolderPath).listItemAllFields();
          console.log(`⚠️⚠️ DEBUG ⚠️⚠️ Folder item after creation:`, folderItem);

          
          // await folderItem.update({
          //   Program: 'General',
          //   ProjectStatus: 'Active',
          //   ProjectNumber: '0000'
          // });


          // -------------------
          console.log(`[ProjectManager] folders.addUsingPath returned:`, newFolder);
          console.log(`[ProjectManager] ✅ Successfully created folder: ${currentPath}`);

          // Add a small delay to avoid rapid API calls
          await new Promise(resolve => setTimeout(resolve, 200)); // 200ms delay between folder operations
        } catch (error: any) {
          console.error(`[ProjectManager] Full error object:`, error);

          // Check if it's a throttling error (406 or 429)
          if (error.status === 406 || error.status === 429 || error.message?.includes('429') || error.message?.includes('throttl')) {
            console.log(`[ProjectManager] Throttling detected, implementing exponential backoff...`);

            // Exponential backoff: 1s, 2s, 4s, 8s
            const backoffDelay = Math.pow(2, i % 4) * 1000;
            console.log(`[ProjectManager] Waiting ${backoffDelay}ms before retrying...`);
            await new Promise(resolve => setTimeout(resolve, backoffDelay));

            // Retry the current folder creation
            i--; // Decrement to retry this folder
            currentPath = parentPath; // Reset current path
            continue;
          }

          // Check if it's a specific SharePoint error
          let specificError = '';
          if (error.data && error.data.responseBody) {
            try {
              const errorBody = JSON.parse(error.data.responseBody);
              if (errorBody.error && errorBody.error.message) {
                specificError = `\nSharePoint Error: ${errorBody.error.message}`;
              }
            } catch (e) {
              // Ignore JSON parse errors
            }
          }

          const errorDetails = `Failed to create folder: ${currentPath}\n` +
            `Parent path: ${parentPath}\n` +
            `Folder name: ${folderName}\n` +
            `Segment ${i + 1} of ${segments.length}\n` +
            `Error: ${error.message || error}${specificError}\n` +
            `Stack: ${error.stack || 'No stack trace available'}`;
          console.error(`[ProjectManager] ❌ ${errorDetails}`);
          throw new Error(errorDetails);
        }
      }

      // All folders created successfully
      console.log(`[ProjectManager] ✅ Complete folder path created successfully: ${folderPath}`);
      return true;
    } catch (error) {
      console.error(`[ProjectManager] ❌ Error creating folder path ${folderPath}:`, error);
      return false;
    }
  }


  private async _updateFileMetadata(fileUrl: string, metadata: { Program?: string; ProjectStatus?: string; ProjectNumber?: string }, listTitle: string, listItemType: string, retryCount: number = 0): Promise<void> {
    try {
      console.log(`[ProjectManager] Updating file metadata for: ${fileUrl}`);
      console.log(`[ProjectManager] Metadata to apply:`, metadata);

      if (!metadata.Program && !metadata.ProjectStatus && !metadata.ProjectNumber) {
        console.log(`[ProjectManager] No metadata to update for file`);
        return;
      }

      // Get the file's list item to get its ID
      const file = this._sp.web.getFileByServerRelativePath(fileUrl);
      let itemData;

      try {
        itemData = await file.listItemAllFields();
      } catch (error: any) {
        // If the file's list item isn't ready yet, wait and retry
        if (retryCount === 0 && (error.status === 404 || error.message?.indexOf('does not exist') > -1)) {
          console.log(`[ProjectManager] File list item not ready yet, waiting 2 seconds...`);
          await new Promise(resolve => setTimeout(resolve, 2000));
          return this._updateFileMetadata(fileUrl, metadata, listTitle, listItemType, 1);
        }
        throw error;
      }

      if (!itemData || !itemData.Id) {
        throw new Error(`No list item found for file: ${fileUrl}`);
      }

      console.log(`[ProjectManager] File list item ID: ${itemData.Id}`);

      // Build the update payload using the provided list item type
      const updatePayload: any = {
        "__metadata": { "type": listItemType }
      };

      if (metadata.Program) updatePayload.Program = metadata.Program;
      if (metadata.ProjectStatus) updatePayload.ProjectStatus = metadata.ProjectStatus;
      if (metadata.ProjectNumber) updatePayload.ProjectNumber = metadata.ProjectNumber;

      // Update the file's metadata using REST API with the provided list title
      const updateEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemData.Id})`;

      const headers: any = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        'odata-version': ''
      };

      const updateResponse = await this.context.spHttpClient.post(
        updateEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: headers,
          body: JSON.stringify(updatePayload)
        }
      );

      if (!updateResponse.ok) {
        throw new Error(`Failed to update file metadata: ${updateResponse.status} ${updateResponse.statusText}`);
      }

      console.log(`[ProjectManager] ✅ Successfully updated file metadata for: ${fileUrl}`);
    } catch (error: any) {
      // Check if it's a throttling error and retry
      if ((error.status === 406 || error.status === 429) && retryCount < 3) {
        const delay = (retryCount + 1) * 2000; // 2s, 4s, 6s
        console.log(`[ProjectManager] File metadata update throttled, retrying after ${delay}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this._updateFileMetadata(fileUrl, metadata, listTitle, listItemType, retryCount + 1);
      }

      console.error(`[ProjectManager] Error updating file metadata:`, error);
      // Continue - metadata update failure shouldn't stop the copy process
    }
  }

  private async _copyFile(sourceFileUrl: string, destFileUrl: string, retryCount: number = 0, onThrottled?: (isThrottled: boolean, retryInfo?: string) => void): Promise<boolean> {
    const MAX_RETRIES = 5;
    const RETRY_DELAYS = [5000, 10000, 20000, 40000, 60000]; // 5s, 10s, 20s, 40s, 60s

    try {
      console.log(`[ProjectManager] Checking if file exists: ${destFileUrl}`);

      // First check if destination file already exists
      try {
        await this._sp.web.getFileByServerRelativePath(destFileUrl).select('Exists')();
        console.log(`[ProjectManager] ⏭️ File already exists, skipping: ${destFileUrl}`);
        return true; // File exists, don't copy
      } catch {
        console.log(`[ProjectManager] File doesn't exist, will copy from: ${sourceFileUrl}`);
      }

      // Check if source file exists
      try {
        await this._sp.web.getFileByServerRelativePath(sourceFileUrl).select('Exists')();
        console.log(`[ProjectManager] ✓ Source file exists: ${sourceFileUrl}`);
      } catch (error: any) {
        // Check if it's a throttling error
        if (error.status === 406 || error.status === 429 || error.message?.indexOf('406') > -1 || error.message?.indexOf('429') > -1) {
          if (retryCount < MAX_RETRIES) {
            const delay = RETRY_DELAYS[retryCount];
            const delaySeconds = Math.round(delay / 1000);
            console.log(`[ProjectManager] ⚠️ Throttling detected (406/429). Retry ${retryCount + 1}/${MAX_RETRIES} after ${delaySeconds} seconds`);

            // Notify about throttling
            if (onThrottled) {
              onThrottled(true, `SharePoint throttling detected. Waiting ${delaySeconds} seconds before retry ${retryCount + 1}/${MAX_RETRIES}...`);
            }

            await new Promise(resolve => setTimeout(resolve, delay));

            // Clear throttling message before retry
            if (onThrottled) {
              onThrottled(false);
            }

            return this._copyFile(sourceFileUrl, destFileUrl, retryCount + 1, onThrottled);
          }
          const errorMsg = `Throttling error (HTTP 406/429) - Too many requests.\n` +
            `SharePoint is temporarily limiting requests. This usually happens with large templates.\n` +
            `Maximum retries (${MAX_RETRIES}) exceeded. Please wait a few minutes and try again.\n` +
            `File: ${sourceFileUrl}`;
          console.error(`[ProjectManager] ❌ ${errorMsg}`);
          throw new Error(errorMsg);
        }

        const errorMsg = `Error accessing source file: ${sourceFileUrl}\n` +
          `Error: ${error.message || error}`;
        console.error(`[ProjectManager] ❌ ${errorMsg}`);
        throw new Error(errorMsg);
      }

      // Copy the file with overwrite flag set to false to prevent duplicates
      console.log(`[ProjectManager] Copying file...`);
      console.log(`[ProjectManager]   From: ${sourceFileUrl}`);
      console.log(`[ProjectManager]   To: ${destFileUrl}`);

      await this._sp.web.getFileByServerRelativePath(sourceFileUrl).copyByPath(destFileUrl, false);
      console.log(`[ProjectManager] ✅ Successfully copied file: ${destFileUrl}`);
      return true;
    } catch (error: any) {
      // Check if it's a throttling error
      if (error.status === 406 || error.status === 429 || error.message?.indexOf('406') > -1 || error.message?.indexOf('429') > -1) {
        if (retryCount < MAX_RETRIES) {
          const delay = RETRY_DELAYS[retryCount];
          const delaySeconds = Math.round(delay / 1000);
          console.log(`[ProjectManager] ⚠️ Throttling detected (406/429). Retry ${retryCount + 1}/${MAX_RETRIES} after ${delaySeconds} seconds`);

          // Notify about throttling
          if (onThrottled) {
            onThrottled(true, `SharePoint throttling detected. Waiting ${delaySeconds} seconds before retry ${retryCount + 1}/${MAX_RETRIES}...`);
          }

          await new Promise(resolve => setTimeout(resolve, delay));

          // Clear throttling message before retry
          if (onThrottled) {
            onThrottled(false);
          }

          return this._copyFile(sourceFileUrl, destFileUrl, retryCount + 1, onThrottled);
        }
        const errorMsg = `Throttling error (HTTP 406/429) - Too many requests.\n` +
          `SharePoint is temporarily limiting requests. This usually happens with large templates.\n` +
          `Maximum retries (${MAX_RETRIES}) exceeded. Please wait a few minutes and try again.\n` +
          `File: ${sourceFileUrl}`;
        throw new Error(errorMsg);
      }

      const errorDetails = `Failed to copy file\n` +
        `Source: ${sourceFileUrl}\n` +
        `Destination: ${destFileUrl}\n` +
        `Error: ${error.message || error}`;
      console.error(`[ProjectManager] ❌ ${errorDetails}`);
      throw new Error(errorDetails);
    }
  }

  public dispose(): void {
    if (this._checkInterval) {
      window.clearInterval(this._checkInterval);
    }
    super.dispose();
  }
}