import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  Stack,
  Text,
  Icon,
  MessageBar,
  MessageBarType,
  IStackTokens,
  IStackStyles,
  Spinner,
  SpinnerSize,
  Dropdown,
  IDropdownOption
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/sites";
import "@pnp/sp/user-custom-actions";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import { IUserCustomActionInfo } from '@pnp/sp/user-custom-actions';
import styles from './TemplateSettingsDialog.module.scss';

export interface ITemplateSettings {
  program: string;
  sourceSite: string;
  sourceLibrary: string;
  sourceFolder: string;
  targetSite: string;
  targetLibrary: string;
  targetFolder: string;
  projectFolderLevels: number;  // 1 or 2 levels deep
  pollingInterval: number;  // Interval in seconds for automatic checking
  copyMode: 'automatic' | 'manual';  // Automatic or Manual mode
}

interface ITemplateSettingsDialogProps {
  sp: SPFI;
  currentSiteUrl: string;
  currentLibrary: string;
  currentFolder: string;
  extensionId: string;
  onSave: (settings: ITemplateSettings) => void;
  onClose: () => void;
  onLoadTemplate?: () => void;  // Optional callback for manual template loading
}

interface ITemplateSettingsDialogState {
  settings: ITemplateSettings;
  programOptions: IDropdownOption[];
  sourceLibraryOptions: IDropdownOption[];
  isLoading: boolean;
  isSaving: boolean;
  error: string;
  successMessage: string;
}

interface IAppCustomizerInfo extends IUserCustomActionInfo {
  ClientSideComponentId: string;
  ClientSideComponentProperties: string;
  Location: string;
}

const stackTokens: IStackTokens = { childrenGap: 15 };
const sectionStackTokens: IStackTokens = { childrenGap: 10 };

const containerStackStyles: IStackStyles = {
  root: {
    padding: '20px 0'
  }
};

const sectionStyles: IStackStyles = {
  root: {
    padding: '15px',
    borderRadius: '4px',
    border: '1px solid #e1e1e1',
    backgroundColor: '#fafafa'
  }
};

class TemplateSettingsDialogContent extends React.Component<ITemplateSettingsDialogProps, ITemplateSettingsDialogState> {
  private siteUrlChangeTimer: number | undefined;

  constructor(props: ITemplateSettingsDialogProps) {
    super(props);

    this.state = {
      settings: {
        program: '',  // Will be populated from dropdown
        sourceSite: '',
        sourceLibrary: 'ProjectDocsTemplates',  // Using internal name, not title
        sourceFolder: '',
        targetSite: props.currentSiteUrl,
        targetLibrary: props.currentLibrary,  // Use the current library where button was clicked
        targetFolder: props.currentFolder,
        projectFolderLevels: 2,  // Default to 2 levels (Location > Project)
        pollingInterval: 8,  // Default to 8 seconds
        copyMode: 'automatic'  // Default to automatic mode
      },
      programOptions: [],
      sourceLibraryOptions: [],
      isLoading: true,
      isSaving: false,
      error: '',
      successMessage: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    await Promise.all([
      this.loadSettings(),
      this.loadProgramChoices(),
      this.loadDocumentLibraries()
    ]);
  }

  public componentWillUnmount(): void {
    // Clear any pending timers
    if (this.siteUrlChangeTimer) {
      clearTimeout(this.siteUrlChangeTimer);
    }
  }

  private async loadProgramChoices(): Promise<void> {
    try {
      console.log('[TemplateSettings] Loading Program field choices...');

      // Get the current library internal name
      const libraryPath = this.props.currentLibrary;
      if (!libraryPath) {
        console.log('[TemplateSettings] No library specified, skipping Program field loading');
        return;
      }

      // Build the full server-relative URL for the library
      const siteUrl = this.props.currentSiteUrl;
      let libraryUrl = '';

      // Extract site server-relative path
      if (siteUrl.indexOf('.sharepoint.com') !== -1) {
        const match = siteUrl.match(/\.sharepoint\.com(.+)/);
        if (match) {
          libraryUrl = `${match[1]}/${libraryPath}`;
        }
      } else {
        // For local development or if URL format is different
        libraryUrl = `${siteUrl}/${libraryPath}`;
      }

      console.log('[TemplateSettings] Getting Program field from library at:', libraryUrl);

      // Use REST API directly to get the field by server-relative URL
      const endpoint = `${this.props.currentSiteUrl}/_api/web/GetList('${libraryUrl}')/fields/getByInternalNameOrTitle('Program')`;

      const response = await fetch(endpoint, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        credentials: 'same-origin'
      });

      if (response.ok) {
        const field = await response.json();

        if (field && field.Choices) {
          const options: IDropdownOption[] = field.Choices.map((choice: string) => ({
            key: choice,
            text: choice
          }));

          console.log('[TemplateSettings] Loaded Program choices:', options);

          this.setState({
            programOptions: options
          });
        }
      } else {
        console.error('[TemplateSettings] Failed to load Program field, status:', response.status);
      }
    } catch (error) {
      console.error('[TemplateSettings] Error loading Program field choices:', error);
      // Don't fail the whole dialog if we can't load program choices
      // User can still use the dialog without program filtering
    }
  }

  private async loadDocumentLibraries(): Promise<void> {
    try {
      console.log('[TemplateSettings] Loading document libraries from current site...');

      // Get all lists from the current site with necessary fields
      const lists = await this.props.sp.web.lists
        .select('Title', 'BaseTemplate', 'Hidden', 'IsCatalog', 'EntityTypeName', 'RootFolder/Name')
        .expand('RootFolder')();

      const libraryOptions = this.filterAndMapLibraries(lists);
      console.log('[TemplateSettings] Loaded document libraries from current site:', libraryOptions);

      // Check if "Project Docs Templates" exists and pre-select it for source library
      let projectDocsTemplates: IDropdownOption | undefined;
      for (let i = 0; i < libraryOptions.length; i++) {
        const lib = libraryOptions[i];
        if (lib.text === 'Project Docs Templates' || lib.key === 'ProjectDocsTemplates') {
          projectDocsTemplates = lib;
          break;
        }
      }

      this.setState(prevState => ({
        sourceLibraryOptions: libraryOptions, // Only for source library dropdown
        settings: {
          ...prevState.settings,
          // Pre-select "Project Docs Templates" if it exists, otherwise keep the default
          sourceLibrary: projectDocsTemplates ? projectDocsTemplates.key as string : prevState.settings.sourceLibrary
        }
      }));
    } catch (error) {
      console.error('[TemplateSettings] Error loading document libraries:', error);
      // Don't fail the whole dialog if we can't load libraries
    }
  }

  private filterAndMapLibraries(lists: any[]): IDropdownOption[] {
    // Filter for document libraries only (BaseTemplate 101)
    // Exclude hidden and system libraries
    const docLibs = lists.filter((list: any) => {
      return list.BaseTemplate === 101 && // Document Library
             !list.Hidden && // Not hidden
             !list.IsCatalog && // Not a catalog
             list.Title !== 'Site Assets' && // Not Site Assets
             list.Title !== 'Site Pages' && // Not Site Pages
             list.Title !== 'Style Library' && // Not Style Library
             list.Title !== 'Form Templates' && // Not Form Templates
             !list.Title.startsWith('_'); // Not system libraries starting with _
    });

    // Create dropdown options using the internal name from RootFolder
    return docLibs.map((lib: any) => {
      // The RootFolder.Name gives us the actual internal name
      const internalName = lib.RootFolder?.Name || lib.Title;
      return {
        key: internalName,
        text: lib.Title, // Display title for user-friendly names
        data: { internalName: internalName, title: lib.Title }
      };
    });
  }

  private async loadDocumentLibrariesFromSite(siteUrl: string): Promise<void> {
    if (!siteUrl) {
      console.log('[TemplateSettings] No site URL provided, skipping library load');
      return;
    }

    try {
      // Normalize the URL by removing trailing slash if present
      const normalizedUrl = siteUrl.replace(/\/$/, '');
      console.log('[TemplateSettings] Loading document libraries from site:', normalizedUrl);

      // Create a new SP context for the specified site using REST API
      const listsEndpoint = `${normalizedUrl}/_api/web/lists?$select=Title,BaseTemplate,Hidden,IsCatalog,EntityTypeName,RootFolder/Name&$expand=RootFolder`;

      const response = await fetch(listsEndpoint, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        credentials: 'same-origin'
      });

      if (!response.ok) {
        throw new Error(`Failed to load libraries from site: ${response.status}`);
      }

      const data = await response.json();
      const lists = data.value || [];

      const libraryOptions = this.filterAndMapLibraries(lists);
      console.log('[TemplateSettings] Loaded libraries from remote site:', libraryOptions);

      // Check if "Project Docs Templates" exists and pre-select it for source library
      let projectDocsTemplates: IDropdownOption | undefined;
      for (let i = 0; i < libraryOptions.length; i++) {
        const lib = libraryOptions[i];
        if (lib.text === 'Project Docs Templates' || lib.key === 'ProjectDocsTemplates') {
          projectDocsTemplates = lib;
          break;
        }
      }

      this.setState(prevState => ({
        sourceLibraryOptions: libraryOptions,
        settings: {
          ...prevState.settings,
          // Pre-select "Project Docs Templates" if it exists
          sourceLibrary: projectDocsTemplates ? projectDocsTemplates.key as string : ''
        },
        error: '' // Clear any previous errors
      }));
    } catch (error) {
      console.error('[TemplateSettings] Error loading libraries from remote site:', error);
      // Clear the source library options if we can't load from the site
      this.setState({
        sourceLibraryOptions: [],
        error: `Failed to load libraries from ${siteUrl}: ${error.message || error}`
      });
    }
  }

  private async loadSettings(): Promise<void> {
    try {
      console.log('[TemplateSettings] Loading settings from UserCustomActions...');

      // Get all user custom actions for the web (not site)
      const siteUserCustomActions = await this.props.sp.web.userCustomActions();

      console.log('[TemplateSettings] All UserCustomActions found:', siteUserCustomActions.length);
      siteUserCustomActions.forEach((uca: any) => {
        console.log('[TemplateSettings] UserCustomAction:', {
          Location: uca.Location,
          ClientSideComponentId: uca.ClientSideComponentId,
          Title: uca.Title
        });
      });

      if (siteUserCustomActions && siteUserCustomActions.length > 0) {
        // Filter for our extension - check multiple possible locations
        const appCustomizers = siteUserCustomActions
          .filter((uca: any) => {
            // Check for various possible location values
            return uca.Location === "ClientSideExtension.ListViewCommandSet.CommandBar" ||
                   uca.Location === "ClientSideExtension.ListViewCommandSet" ||
                   uca.Location === "CommandBar";
          })
          .filter((uca: IAppCustomizerInfo) => uca.ClientSideComponentId === this.props.extensionId);

        console.log('[TemplateSettings] Filtered extensions found:', appCustomizers.length);

        if (appCustomizers.length > 0) {
          const uca = appCustomizers[0] as IAppCustomizerInfo;

          if (uca.ClientSideComponentProperties) {
            try {
              const props = JSON.parse(uca.ClientSideComponentProperties);

              if (props.templateSettings) {
                console.log('[TemplateSettings] Found existing settings:', props.templateSettings);

                this.setState({
                  settings: {
                    ...this.state.settings,
                    ...props.templateSettings
                  },
                  isLoading: false
                });
                return;
              }
            } catch (parseError) {
              console.error('[TemplateSettings] Error parsing properties:', parseError);
            }
          }
        }
      }

      console.log('[TemplateSettings] No existing settings found, using defaults');

      // Set default source site to current site's parent if it has one
      const defaultSourceSite = this.props.currentSiteUrl.replace('/ProjectsRepository-DEV', '/ProjectsRepository');

      this.setState({
        settings: {
          ...this.state.settings,
          sourceSite: defaultSourceSite
        },
        isLoading: false
      });

    } catch (error) {
      console.error('[TemplateSettings] Error loading settings:', error);
      this.setState({
        error: 'Failed to load settings: ' + error.message,
        isLoading: false
      });
    }
  }

  private async ensureProjectCreationLogsList(): Promise<void> {
    try {
      console.log('[TemplateSettings] Ensuring ProjectCreationLogs list exists...');

      const listTitle = 'ProjectCreationLogs';

      // Check if list already exists
      try {
        await this.props.sp.web.lists.getByTitle(listTitle)();
        console.log('[TemplateSettings] ProjectCreationLogs list already exists');
        return;
      } catch (error) {
        // List doesn't exist, create it
        console.log('[TemplateSettings] ProjectCreationLogs list not found, creating...');
      }

      // Create the list
      const listCreationInfo = {
        Title: listTitle,
        Description: 'Logs for template project creation',
        BaseTemplate: 100, // Generic List
        AllowContentTypes: false,
        ContentTypesEnabled: false
      };

      const createdList = await this.props.sp.web.lists.add(listTitle, listCreationInfo.Description, listCreationInfo.BaseTemplate, false);
      console.log('[TemplateSettings] List created:', createdList.data.Title);

      // Add custom fields
      const list = this.props.sp.web.lists.getByTitle(listTitle);

      // Add Program field (Single line of text)
      await list.fields.addText('Program', {
        Description: 'Program name',
        Required: false,
        MaxLength: 255
      });

      // Add Status field (Single line of text)
      await list.fields.addText('Status', {
        Description: 'Creation status (Success/Failed)',
        Required: true,
        MaxLength: 50
      });

      // Add ItemsCopied field (Number)
      await list.fields.addNumber('ItemsCopied', {
        Description: 'Number of items copied',
        MinimumValue: 0
      });

      // Add timestamp field for tracking
      await list.fields.addDateTime('CreatedDateTime', {
        Description: 'When the project was created',
        DisplayFormat: 1, // DateTime
        FriendlyDisplayFormat: 1
      });

      // Add error details field for failed creations
      await list.fields.addMultilineText('ErrorDetails', {
        Description: 'Error details if creation failed',
        NumberOfLines: 10,
        RichText: false
      });

      // Update default view to include new fields
      const defaultView = list.defaultView;
      await defaultView.fields.add('Program');
      await defaultView.fields.add('Status');
      await defaultView.fields.add('ItemsCopied');
      await defaultView.fields.add('CreatedDateTime');

      console.log('[TemplateSettings] ProjectCreationLogs list created successfully with all fields');

    } catch (error) {
      console.error('[TemplateSettings] Error ensuring ProjectCreationLogs list:', error);
      // Don't throw - logging is not critical for the extension to work
    }
  }

  private async saveSettings(): Promise<void> {
    this.setState({ isSaving: true, error: '', successMessage: '' });

    try {
      console.log('[TemplateSettings] Saving settings:', this.state.settings);

      // Ensure the logging list exists
      await this.ensureProjectCreationLogsList();

      // Get all user custom actions from web (not site)
      const siteUserCustomActions = await this.props.sp.web.userCustomActions();

      console.log('[TemplateSettings] Saving - All UserCustomActions found:', siteUserCustomActions.length);

      if (!siteUserCustomActions || siteUserCustomActions.length === 0) {
        throw new Error('No UserCustomActions found. Please ensure the extension is properly deployed to this site.');
      }

      // Find our extension - check multiple possible locations
      const appCustomizers = siteUserCustomActions
        .filter((uca: any) => {
            // Check for various possible location values
            return uca.Location === "ClientSideExtension.ListViewCommandSet.CommandBar" ||
                   uca.Location === "ClientSideExtension.ListViewCommandSet" ||
                   uca.Location === "CommandBar";
          })
        .filter((uca: IAppCustomizerInfo) => uca.ClientSideComponentId === this.props.extensionId);

      console.log('[TemplateSettings] Saving - Filtered extensions found:', appCustomizers.length);

      if (appCustomizers.length === 0) {
        // Log all UserCustomActions for debugging
        console.error('[TemplateSettings] Could not find extension. All UserCustomActions:');
        siteUserCustomActions.forEach((uca: any) => {
          console.error('[TemplateSettings] - Location:', uca.Location, 'ID:', uca.ClientSideComponentId);
        });
        throw new Error(`Template Manager extension not found. Looking for ID: ${this.props.extensionId}`);
      }

      const uca = appCustomizers[0] as IAppCustomizerInfo;

      // Parse existing properties or create new object
      let existingProps = {};
      if (uca.ClientSideComponentProperties) {
        try {
          existingProps = JSON.parse(uca.ClientSideComponentProperties);
        } catch (e) {
          console.warn('[TemplateSettings] Could not parse existing properties, starting fresh');
        }
      }

      // Update properties with new settings
      const updatedProps = {
        ...existingProps,
        templateSettings: this.state.settings,
        isConfigured: true
      };

      console.log('[TemplateSettings] Updating UserCustomAction with ID:', uca.Id);

      // Update the user custom action - cast to any to handle property name
      const userCustomAction = this.props.sp.web.userCustomActions.getById(uca.Id);
      const updateProps: any = {
        ClientSideComponentProperties: JSON.stringify(updatedProps)
      };
      await userCustomAction.update(updateProps);

      console.log('[TemplateSettings] Settings saved successfully');

      this.setState({
        isSaving: false,
        successMessage: 'Settings saved successfully!'
      });

      // Call parent callback
      this.props.onSave(this.state.settings);

      // Close dialog after short delay
      setTimeout(() => {
        this.props.onClose();
      }, 1500);

    } catch (error) {
      console.error('[TemplateSettings] Error saving settings:', error);
      this.setState({
        isSaving: false,
        error: 'Failed to save settings: ' + error.message
      });
    }
  }

  private handleInputChange = (field: keyof ITemplateSettings) => (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({
      settings: {
        ...this.state.settings,
        [field]: newValue || ''
      }
    });
  }

  private handleSourceSiteChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    const siteUrl = newValue || '';
    console.log('[TemplateSettings] Source site URL changed to:', siteUrl);

    this.setState({
      settings: {
        ...this.state.settings,
        sourceSite: siteUrl
      }
    });

    // Clear the previous timer if exists
    if (this.siteUrlChangeTimer) {
      clearTimeout(this.siteUrlChangeTimer);
    }

    // If the URL looks complete (ends with a site name), load libraries after a short delay
    if (siteUrl && this.isValidSiteUrl(siteUrl)) {
      console.log('[TemplateSettings] Valid site URL detected, scheduling library load...');
      this.siteUrlChangeTimer = window.setTimeout(async () => {
        console.log('[TemplateSettings] Timer triggered, loading libraries from:', siteUrl);
        await this.loadDocumentLibrariesFromSite(siteUrl);
      }, 500); // Wait 500ms after user stops typing
    } else if (siteUrl) {
      console.log('[TemplateSettings] URL not valid yet:', siteUrl);
    }
  }

  private isValidSiteUrl(url: string): boolean {
    // Check if URL looks complete (has protocol and at least one path segment after domain)
    // Allow optional trailing slash
    const urlPattern = /^https?:\/\/[^\/]+\.sharepoint\.com\/sites\/[^\/]+\/?$/i;
    return urlPattern.test(url);
  }

  private handleProgramChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({
        settings: {
          ...this.state.settings,
          program: option.key as string
        }
      });
    }
  }

  private handleProjectStructureChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({
        settings: {
          ...this.state.settings,
          projectFolderLevels: option.key as number
        }
      });
    }
  }

  private handleCopyModeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({
        settings: {
          ...this.state.settings,
          copyMode: option.key as ('automatic' | 'manual')
        }
      });
    }
  }

  private handlePollingIntervalChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const interval = parseInt(newValue || '8', 10);
    // Ensure interval is between 1 and 60 seconds
    const validInterval = Math.min(60, Math.max(1, interval));

    this.setState({
      settings: {
        ...this.state.settings,
        pollingInterval: validInterval
      }
    });
  }

  private handleLoadTemplate = async (): Promise<void> => {
    // Save settings first
    await this.handleSave();

    // Then trigger the template loading
    if (this.props.onLoadTemplate) {
      this.props.onLoadTemplate();
    }
  }

  private handleSave = async (): Promise<void> => {
    // Validate required fields
    const { settings } = this.state;

    if (!settings.sourceSite || !settings.sourceLibrary) {
      this.setState({
        error: 'Source site and library are required.'
      });
      return;
    }

    if (!settings.targetSite || !settings.targetLibrary) {
      this.setState({
        error: 'Target site and library are required.'
      });
      return;
    }

    await this.saveSettings();
  }

  public render(): React.ReactElement<ITemplateSettingsDialogProps> {
    const { settings, programOptions, isLoading, isSaving, error, successMessage } = this.state;

    if (isLoading) {
      return (
        <Dialog
          hidden={false}
          onDismiss={this.props.onClose}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: 'Template Settings',
            subText: 'Loading configuration...'
          }}
          modalProps={{
            isBlocking: true,
            styles: {
              main: {
                minWidth: 800,
                maxWidth: '90vw',
                width: 900,
                margin: 'auto'
              }
            },
            containerClassName: styles.dialogContainer
          }}
        >
          <Stack tokens={stackTokens} horizontalAlign="center" styles={{ root: { padding: '40px 0' } }}>
            <Spinner size={SpinnerSize.large} label="Loading settings..." />
          </Stack>
        </Dialog>
      );
    }

    return (
      <Dialog
        hidden={false}
        onDismiss={this.props.onClose}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Template Settings',
          subText: 'Configure where templates are copied from and to when loading templates into project folders.',
         }}
        modalProps={{
          isBlocking: true,
          styles: {
            main: {
              minWidth: 800,
              maxWidth: '90vw',
              width: 900,
              margin: 'auto'
            }
          },
          containerClassName: styles.dialogContainer
        }}
      >
        <Stack tokens={stackTokens} styles={containerStackStyles}>
            {error && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={true}
                onDismiss={() => this.setState({ error: '' })}
              >
                {error}
              </MessageBar>
            )}

            {successMessage && (
              <MessageBar
                messageBarType={MessageBarType.success}
                onDismiss={() => this.setState({ successMessage: '' })}
              >
                {successMessage}
              </MessageBar>
            )}

            {/* Program Selection */}
            <Stack tokens={sectionStackTokens} styles={sectionStyles}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <Icon iconName="Tag" styles={{ root: { fontSize: 20, color: '#5c2d91' } }} />
                <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                  Program Filter
                </Text>
              </Stack>

              <Text variant="small" styles={{ root: { color: '#666', marginBottom: 10 } }}>
                Select a program to filter templates. Only templates marked with this program (or General) will be loaded.
              </Text>

              <Dropdown
                label="Select Program"
                selectedKey={settings.program}
                onChange={this.handleProgramChange}
                placeholder="Select a program"
                options={programOptions}
                disabled={isSaving || programOptions.length === 0}
                errorMessage={programOptions.length === 0 ? "No programs available in the library" : undefined}
              />
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 20 }}>
              {/* Source Configuration */}
              <Stack.Item grow={1}>
                <Stack tokens={sectionStackTokens} styles={sectionStyles}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                    <Icon iconName="FolderOpen" styles={{ root: { fontSize: 20, color: '#0078d4' } }} />
                    <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                      Source (Templates)
                    </Text>
                  </Stack>

                  <Text variant="small" styles={{ root: { color: '#666' } }}>
                    Specify where template folders and files are stored
                  </Text>

                  <TextField
                    label="Site URL"
                    value={settings.sourceSite}
                    onChange={this.handleSourceSiteChange}
                    placeholder="https://tenant.sharepoint.com/sites/Templates"
                    required
                    disabled={isSaving}
                  />

                  <Dropdown
                    label="Document Library"
                    selectedKey={settings.sourceLibrary}
                    onChange={(event, option) => {
                      if (option) {
                        this.setState(prevState => ({
                          settings: {
                            ...prevState.settings,
                            sourceLibrary: option.key as string
                          }
                        }));
                      }
                    }}
                    options={this.state.sourceLibraryOptions}
                    placeholder="Select a document library"
                    required
                    disabled={isSaving}
                  />
                  <Text variant="small" styles={{ root: { color: '#666', marginTop: -8 } }}>
                    Select the document library containing templates
                  </Text>

                  <TextField
                    label="Folder Path (optional)"
                    value={settings.sourceFolder}
                    onChange={this.handleInputChange('sourceFolder')}
                    placeholder="Leave empty for root or enter path like 'Folder1/Folder2'"
                    disabled={isSaving}
                  />
                </Stack>
              </Stack.Item>

              {/* Arrow */}
              <Stack.Item align="center" styles={{ root: { paddingTop: 60 } }}>
                <Icon
                  iconName="DoubleChevronRight"
                  styles={{
                    root: {
                      fontSize: 24,
                      color: '#0078d4',
                      animation: 'pulse 2s infinite'
                    }
                  }}
                />
              </Stack.Item>

              {/* Target Configuration */}
              <Stack.Item grow={1}>
                <Stack tokens={sectionStackTokens} styles={sectionStyles}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                    <Icon iconName="FolderHorizontal" styles={{ root: { fontSize: 20, color: '#107c10' } }} />
                    <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                      Target (Projects)
                    </Text>
                  </Stack>

                  <Text variant="small" styles={{ root: { color: '#666' } }}>
                    Where templates will be copied to (automatically determined by current location)
                  </Text>

                  <TextField
                    label="Site URL"
                    value={settings.targetSite}
                    onChange={this.handleInputChange('targetSite')}
                    placeholder="https://tenant.sharepoint.com/sites/Projects"
                    required
                    disabled={true}  // Always disabled - auto-populated from current context
                    description="Auto-populated from current location"
                  />

                  <TextField
                    label="Document Library"
                    value={settings.targetLibrary}
                    onChange={this.handleInputChange('targetLibrary')}
                    placeholder="Document library internal name"
                    required
                    disabled={true}  // Always disabled - auto-populated from current context
                    description="Auto-populated from current location (internal name)"
                  />

                  <Dropdown
                    label="Project Structure"
                    selectedKey={settings.projectFolderLevels}
                    onChange={this.handleProjectStructureChange}
                    options={[
                      { key: 1, text: 'Customer Folder (1 level)' },
                      { key: 2, text: 'Location Folder > Project Folder (2 levels)' }
                    ]}
                    disabled={isSaving}
                  />
                  <Text variant="small" styles={{ root: { color: '#666', marginTop: -8 } }}>
                    Define the folder hierarchy where templates should be loaded
                  </Text>
                </Stack>
              </Stack.Item>
            </Stack>

            {/* Copy Mode and Polling Settings */}
            <Stack tokens={sectionStackTokens} styles={sectionStyles}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <Icon iconName="Settings" styles={{ root: { fontSize: 20, color: '#5c2d91' } }} />
                <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                  Copy Settings
                </Text>
              </Stack>

              <Text variant="small" styles={{ root: { color: '#666', marginBottom: 10 } }}>
                Configure how and when templates are copied to project folders.
              </Text>

              <Dropdown
                label="Copy Mode"
                selectedKey={settings.copyMode}
                onChange={this.handleCopyModeChange}
                options={[
                  { key: 'automatic', text: 'Automatic - Copy when navigating to project folders' },
                  { key: 'manual', text: 'Manual - Copy only when Load Template button is clicked' }
                ]}
                disabled={isSaving}
              />
              <Text variant="small" styles={{ root: { color: '#666', marginTop: -8 } }}>
                Choose whether templates are loaded automatically or manually
              </Text>

              {settings.copyMode === 'automatic' && (
                <TextField
                  label="Check Interval (seconds)"
                  type="number"
                  min={1}
                  max={60}
                  value={settings.pollingInterval.toString()}
                  onChange={this.handlePollingIntervalChange}
                  disabled={isSaving}
                  description="How often to check the current folder (1-60 seconds)"
                />
              )}
            </Stack>

            <MessageBar messageBarType={MessageBarType.info}>
              <Text variant="small">
                <strong>Note:</strong> These settings define where templates come from and at what folder level they should be loaded.
                The template loader will only trigger when you navigate to a folder at the specified project structure level.
              </Text>
            </MessageBar>
          </Stack>

        <DialogFooter>
          <PrimaryButton
            onClick={this.handleSave}
            text="Save Settings"
            disabled={isSaving}
          />
          {settings.copyMode === 'manual' && this.props.onLoadTemplate && (
            <DefaultButton
              onClick={this.handleLoadTemplate}
              text="Load Template"
              disabled={isSaving}
              iconProps={{ iconName: 'Download' }}
              styles={{ root: { marginLeft: 8 } }}
            />
          )}
          <DefaultButton
            onClick={this.props.onClose}
            text="Cancel"
            disabled={isSaving}
          />
        </DialogFooter>
      </Dialog>
    );
  }
}

export class TemplateSettingsDialog extends BaseDialog {
  private sp: SPFI;
  private currentSiteUrl: string;
  private currentLibrary: string;
  private currentFolder: string;
  private extensionId: string;
  private onSave: (settings: ITemplateSettings) => void;
  private onLoadTemplate?: () => void;

  constructor(
    sp: SPFI,
    currentSiteUrl: string,
    currentLibrary: string,
    currentFolder: string,
    extensionId: string,
    onSave?: (settings: ITemplateSettings) => void,
    onLoadTemplate?: () => void
  ) {
    super();
    this.sp = sp;
    this.currentSiteUrl = currentSiteUrl;
    this.currentLibrary = currentLibrary;
    this.currentFolder = currentFolder;
    this.extensionId = extensionId;
    this.onSave = onSave || (() => {});
    this.onLoadTemplate = onLoadTemplate;
  }

  public render(): void {
    ReactDOM.render(
      <TemplateSettingsDialogContent
        sp={this.sp}
        currentSiteUrl={this.currentSiteUrl}
        currentLibrary={this.currentLibrary}
        currentFolder={this.currentFolder}
        extensionId={this.extensionId}
        onSave={this.onSave}
        onClose={() => this.close()}
        onLoadTemplate={this.onLoadTemplate}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}