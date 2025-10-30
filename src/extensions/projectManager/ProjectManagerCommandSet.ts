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
import TemplateLoaderDialog from './TemplateLoaderDialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProjectManagerCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ProjectManagerCommandSet';

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

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ProjectManagerCommandSet');
    console.log('[ProjectManager] Extension initialized successfully');

    // Initialize PnP
    this._sp = spfi().using(SPFx(this.context));

    // Start checking for empty folders
    this._startEmptyFolderCheck();
    console.log('[ProjectManager] Started empty folder monitoring');

    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    // Remove the CREATE_PROJECT command as it's no longer needed
    // The extension will now work automatically on empty folders
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    // No manual commands to execute
  }

  private _startEmptyFolderCheck(): void {
    // Check every 2 seconds
    this._checkInterval = window.setInterval(() => {
      this._checkCurrentFolder();
    }, 2000);
  }

  private async _checkCurrentFolder(): Promise<void> {
    try {
      // Get current location info
      const pageContext = this.context.pageContext;
      if (!pageContext || !pageContext.list) {
        return;
      }

      const listTitle = pageContext.list.title;
      console.log('[ProjectManager] Checking folder in library:', listTitle);

      // Get folder URL from the page URL
      let folderUrl = '';
      const url = window.location.href;
      console.log('[ProjectManager] Current URL:', url);

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

      console.log('[ProjectManager] Extracted folder URL:', folderUrl);

      // Check if we're in the Documents library
      if (listTitle !== 'Documents') {
        console.log('[ProjectManager] Not in Documents library, skipping');
        return;
      }

      // Parse the folder path to check if we're at the right level
      // Expected pattern: /sites/ProjectsRepository/Shared Documents/Projects/[Program]/[LocationSubfolder]/[ProjectSubfolder]
      const pathSegments = folderUrl.split('/').filter((s: string) => s);
      console.log('[ProjectManager] Path segments:', pathSegments);

      // Find the index of "Documents" or "Shared Documents" in the path
      let docsIndex = -1;
      for (let i = 0; i < pathSegments.length; i++) {
        if (pathSegments[i] === 'Documents' || pathSegments[i] === 'Shared Documents') {
          docsIndex = i;
          break;
        }
      }
      if (docsIndex === -1) {
        console.log('[ProjectManager] Could not find Documents or Shared Documents in path');
        return;
      }

      // Check if we have the correct path structure
      const relativeSegments = pathSegments.slice(docsIndex + 1);
      console.log('[ProjectManager] Segments after Documents:', relativeSegments);

      // We need exactly 4 segments after Documents: Projects/[Program]/[LocationSubfolder]/[ProjectSubfolder]
      if (relativeSegments.length !== 4) {
        console.log('[ProjectManager] Wrong folder depth. Expected 4 segments, got:', relativeSegments.length);
        return;
      }

      // First segment must be "Projects"
      if (relativeSegments[0] !== 'Projects') {
        console.log('[ProjectManager] First segment is not "Projects":', relativeSegments[0]);
        return;
      }

      // We're at the project subfolder level (one level deeper now)
      const programName = relativeSegments[1];
      const locationName = relativeSegments[2];
      const projectName = relativeSegments[3];
      console.log('[ProjectManager] Detected folder structure - Program:', programName, 'Location:', locationName, 'Project:', projectName);

      // Check if we're already loading a template
      if (this._isLoadingTemplate) {
        console.log('[ProjectManager] Template loading already in progress, skipping check');
        return;
      }

      // Check if dialog has already been shown for this folder
      const dialogKey = `templateDialog_${folderUrl}`;
      const dialogShown = sessionStorage.getItem(dialogKey);
      const templateLoaded = sessionStorage.getItem(`${dialogKey}_loaded`);

      if (dialogShown || templateLoaded) {
        console.log('[ProjectManager] Dialog already shown or template already loaded for this folder');
        return;
      }

      // Check if folder is empty
      const isEmpty = await this._isFolderEmpty(folderUrl);
      console.log('[ProjectManager] Folder empty check result:', isEmpty);

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
      // Check for both files and folders
      const filesEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Files`;
      const foldersEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Folders`;

      const [filesResponse, foldersResponse] = await Promise.all([
        this.context.spHttpClient.get(filesEndpoint, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(foldersEndpoint, SPHttpClient.configurations.v1)
      ]);

      if (!filesResponse.ok || !foldersResponse.ok) {
        console.log('[ProjectManager] Error checking folder contents, treating as non-empty');
        return false;
      }

      const filesData = await filesResponse.json();
      const foldersData = await foldersResponse.json();

      // Filter out Forms folder if it exists
      const folders = foldersData.value.filter((f: any) => f.Name !== 'Forms');

      // Log what we found
      console.log(`[ProjectManager] Folder contents - Files: ${filesData.value.length}, Folders: ${folders.length}`);
      if (folders.length > 0) {
        console.log('[ProjectManager] Existing folders:', folders.map((f: any) => f.Name).join(', '));
      }

      return filesData.value.length === 0 && folders.length === 0;
    } catch (error) {
      console.error('[ProjectManager] Error checking if folder is empty:', error);
      return false;
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

    // Show the template loader dialog
    const dialog = new TemplateLoaderDialog(
      folderUrl,
      programName,
      locationName,
      projectName,
      async (destFolder: string, onProgress?: (percent: number, message: string) => void) => {
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
      }
    );

    try {
      await dialog.show();
    } catch (error) {
      console.error('[ProjectManager] Error showing dialog:', error);
      this._isLoadingTemplate = false;
      sessionStorage.removeItem(dialogKey);
    }
  }

  private async _loadTemplate(destinationFolderUrl: string, currentProgramName: string, onProgress?: (percent: number, message: string) => void): Promise<void> {
    try {
      console.log('[ProjectManager] Loading template for program:', currentProgramName);
      console.log('[ProjectManager] Destination folder:', destinationFolderUrl);

      // Source is the Templates library root
      const sourceUrl = '/sites/ProjectsRepository/Templates';

      // Get all items from Templates library with metadata
      const items = await this._getAllItemsFromFolderWithMetadata(sourceUrl);

      console.log('[ProjectManager] Retrieved items from Templates library:');
      console.log('[ProjectManager] Total folders found:', items.folders.length);
      console.log('[ProjectManager] Folder list:', items.folders.map(f => `${f.Name} (${f.ServerRelativeUrl})`));

      if (items.files.length === 0 && items.folders.length === 0) {
        throw new Error('No template files found in the Templates library');
      }

      // Filter items based on Program field
      const filteredFolders = items.folders.filter(folder => {
        const programValue = folder.Program || '';
        console.log(`[ProjectManager] Evaluating folder "${folder.Name}" at path "${folder.ServerRelativeUrl}" with Program value: "${programValue}"`);

        // Copy if Program is empty, "General", or matches current program
        return !programValue ||
               programValue === '' ||
               programValue.toLowerCase() === 'general' ||
               programValue === currentProgramName;
      });

      const filteredFiles = items.files.filter(file => {
        const programValue = file.Program || '';
        console.log(`[ProjectManager] File "${file.Name}" at "${file.ServerRelativeUrl}" has Program value: "${programValue}"`);

        // Copy if Program is empty, "General", or matches current program
        return !programValue ||
               programValue === '' ||
               programValue.toLowerCase() === 'general' ||
               programValue === currentProgramName;
      });

      console.log(`[ProjectManager] Filtered folders before deduplication: ${filteredFolders.length}`);
      console.log('[ProjectManager] Filtered folder names:', filteredFolders.map(f => f.Name));

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
        if (onProgress) {
          onProgress(Math.round((processedItems / totalItems) * 100), `Creating folder: ${folder.Name}`);
        }

        const relativePath = folder.ServerRelativeUrl.replace(sourceUrl, '');
        const destFolderUrl = `${destinationFolderUrl}${relativePath}`;

        console.log(`[ProjectManager] [${foldersCreated + foldersFailed + 1}/${uniqueFolders.length}] Processing folder: "${folder.Name}"`);

        const success = await this._createFolderPath(destFolderUrl);
        if (!success) {
          // STOP IMMEDIATELY on failure
          const errorMsg = `Failed to create folder: "${folder.Name}" at path: ${destFolderUrl}`;
          console.error(`[ProjectManager] ❌ CRITICAL ERROR: ${errorMsg}`);
          throw new Error(errorMsg);
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

      for (const file of filteredFiles) {
        if (onProgress) {
          onProgress(Math.round((processedItems / totalItems) * 100), `Copying file: ${file.Name}`);
        }

        const relativePath = file.ServerRelativeUrl.replace(sourceUrl, '');
        const destFileUrl = `${destinationFolderUrl}${relativePath}`;

        // Double-check and ensure parent folder exists before copying file
        const parentFolderPath = destFileUrl.substring(0, destFileUrl.lastIndexOf('/'));
        console.log(`[ProjectManager] Ensuring parent folder exists for file: ${parentFolderPath}`);

        const folderCreated = await this._ensureFolderExists(parentFolderPath);
        if (!folderCreated) {
          // STOP IMMEDIATELY on folder creation failure
          const errorMsg = `Failed to ensure parent folder for file "${file.Name}". Parent folder path: ${parentFolderPath}`;
          console.error(`[ProjectManager] ❌ CRITICAL ERROR: ${errorMsg}`);
          throw new Error(errorMsg);
        }

        console.log(`[ProjectManager] [${filesCopied + filesFailed + 1}/${filteredFiles.length}] Processing file: "${file.Name}"`);

        const success = await this._copyFile(file.ServerRelativeUrl, destFileUrl);
        if (!success) {
          // STOP IMMEDIATELY on file copy failure
          const errorMsg = `Failed to copy file: "${file.Name}" from ${file.ServerRelativeUrl} to ${destFileUrl}`;
          console.error(`[ProjectManager] ❌ CRITICAL ERROR: ${errorMsg}`);
          throw new Error(errorMsg);
        }
        filesCopied++;
        processedItems++;
      }

      console.log(`[ProjectManager] File copying complete: ${filesCopied} files copied successfully`);

      // Final summary
      console.log('[ProjectManager] ========== TEMPLATE LOADING SUMMARY ==========');
      console.log(`[ProjectManager] ✅ Successfully created ${foldersCreated} folders`);
      console.log(`[ProjectManager] ✅ Successfully copied ${filesCopied} files`);
      console.log(`[ProjectManager] ✅ All operations completed successfully!`);

      if (onProgress) {
        onProgress(100, 'Template loaded successfully!');
      }

    } catch (error) {
      console.error('Error loading template:', error);
      throw error;
    }
  }

  private async _getAllItemsFromFolderWithMetadata(folderUrl: string, visitedPaths?: { [key: string]: boolean }): Promise<{ files: IFileItem[], folders: IFolderItem[] }> {
    try {
      // Initialize visited paths object if not provided
      if (!visitedPaths) {
        visitedPaths = {};
      }

      // Prevent infinite loops and duplicate processing
      if (visitedPaths[folderUrl]) {
        console.log(`[ProjectManager] Skipping already visited folder: ${folderUrl}`);
        return { files: [], folders: [] };
      }
      visitedPaths[folderUrl] = true;

      // Get all files and folders with Program field
      const filesEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Files?$select=*,ListItemAllFields/Program&$expand=ListItemAllFields`;
      const foldersEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Folders?$select=*,ListItemAllFields/Program&$expand=ListItemAllFields`;

      const [filesResponse, foldersResponse] = await Promise.all([
        this.context.spHttpClient.get(filesEndpoint, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(foldersEndpoint, SPHttpClient.configurations.v1)
      ]);

      if (!filesResponse.ok || !foldersResponse.ok) {
        throw new Error('Failed to get folder contents with metadata');
      }

      const filesData = await filesResponse.json();
      const foldersData = await foldersResponse.json();

      let allFiles: IFileItem[] = filesData.value.map((f: any) => ({
        Name: f.Name,
        ServerRelativeUrl: f.ServerRelativeUrl,
        TimeCreated: f.TimeCreated,
        TimeLastModified: f.TimeLastModified,
        Length: f.Length,
        UniqueId: f.UniqueId,
        Program: f.ListItemAllFields ? f.ListItemAllFields.Program : undefined
      }));

      let allFolders: IFolderItem[] = foldersData.value
        .filter((f: any) => f.Name !== 'Forms')
        .map((f: any) => ({
          Name: f.Name,
          ServerRelativeUrl: f.ServerRelativeUrl,
          TimeCreated: f.TimeCreated,
          TimeLastModified: f.TimeLastModified,
          ItemCount: f.ItemCount,
          UniqueId: f.UniqueId,
          Program: f.ListItemAllFields ? f.ListItemAllFields.Program : undefined
        }));

      // Recursively get items from subfolders
      for (const folder of allFolders) {
        const subItems = await this._getAllItemsFromFolderWithMetadata(folder.ServerRelativeUrl, visitedPaths);
        allFiles = allFiles.concat(subItems.files);
        allFolders = allFolders.concat(subItems.folders);
      }

      return { files: allFiles, folders: allFolders };
    } catch (error) {
      console.error('Error getting folder items with metadata:', error);
      return { files: [], folders: [] };
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
          console.log(`[ProjectManager] folders.addUsingPath returned:`, newFolder);
          console.log(`[ProjectManager] ✅ Successfully created folder: ${currentPath}`);

          // Verify the folder was actually created
          console.log(`[ProjectManager] Verifying folder creation...`);
          const verifyFolder = await this._sp.web.getFolderByServerRelativePath(currentPath).select('Exists')();
          if (!verifyFolder.Exists) {
            const errorMsg = `Folder creation verification failed: ${currentPath}\n` +
              `The folder was supposedly created but doesn't exist when checked`;
            console.error(`[ProjectManager] ❌ ${errorMsg}`);
            throw new Error(errorMsg);
          }
          console.log(`[ProjectManager] ✅ Folder verified to exist: ${currentPath}`);
        } catch (error) {
          console.error(`[ProjectManager] Full error object:`, error);

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

      // Final verification that the complete path exists
      try {
        const finalCheck = await this._sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
        if (finalCheck.Exists) {
          console.log(`[ProjectManager] ✅ Final verification: Folder path created successfully: ${folderPath}`);
          return true;
        } else {
          console.error(`[ProjectManager] ❌ Final verification failed: ${folderPath}`);
          return false;
        }
      } catch {
        console.error(`[ProjectManager] ❌ Could not verify final folder path: ${folderPath}`);
        return false;
      }
    } catch (error) {
      console.error(`[ProjectManager] ❌ Error creating folder path ${folderPath}:`, error);
      return false;
    }
  }


  private async _copyFile(sourceFileUrl: string, destFileUrl: string): Promise<boolean> {
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
      } catch (error) {
        const errorMsg = `Source file doesn't exist: ${sourceFileUrl}\n` +
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
    } catch (error) {
      const errorDetails = `Failed to copy file\n` +
        `Source: ${sourceFileUrl}\n` +
        `Destination: ${destFileUrl}\n` +
        `Error: ${error.message || error}\n` +
        `Stack: ${error.stack || 'No stack trace available'}`;
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