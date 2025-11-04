import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import {
  Dialog,
  DialogFooter,
  DialogContent,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  Stack,
  Text,
  FontWeights,
  IStackTokens,
  DialogType,
  Icon
} from 'office-ui-fabric-react';

export interface ITemplateLoaderDialogProps {
  folderUrl: string;
  programName: string;
  locationName: string;
  projectName: string;
  loadTemplate: (folderUrl: string, onProgress?: (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => void) => Promise<void>;
  onCancel: () => void;
  onClose: () => void;
  autoStart?: boolean;  // Auto-start loading when in automatic mode
}

interface ITemplateLoaderDialogState {
  isLoading: boolean;
  loadingProgress: number;
  loadingMessage: string;
  currentItem: number;
  totalItems: number;
  isComplete: boolean;
  error: string;
  showDialog: boolean;
  isThrottled: boolean;
}

const stackTokens: IStackTokens = { childrenGap: 20 };

class TemplateLoaderDialogContent extends React.Component<ITemplateLoaderDialogProps, ITemplateLoaderDialogState> {
  constructor(props: ITemplateLoaderDialogProps) {
    super(props);
    this.state = {
      isLoading: false,
      loadingProgress: 0,
      loadingMessage: '',
      currentItem: 0,
      totalItems: 0,
      isComplete: false,
      error: '',
      showDialog: true,
      isThrottled: false
    };
  }

  public componentDidMount(): void {
    // Add event listener to prevent closing during loading
    window.addEventListener('beforeunload', this._handleBeforeUnload);

    // Auto-start loading if autoStart is true
    if (this.props.autoStart) {
      // Start loading after a brief delay to ensure dialog is rendered
      setTimeout(() => {
        this._handleLoadTemplate();
      }, 500);
    }
  }

  public componentWillUnmount(): void {
    // Remove event listener when component unmounts
    window.removeEventListener('beforeunload', this._handleBeforeUnload);
  }

  private _handleBeforeUnload = (e: BeforeUnloadEvent): BeforeUnloadEvent | undefined => {
    // Only prevent closing if we're currently loading
    if (this.state.isLoading) {
      e.preventDefault();
      e.returnValue = 'Template is currently being loaded. Are you sure you want to leave?';
      return e;
    }
    return undefined;
  }

  public render(): JSX.Element {
    const { locationName, projectName } = this.props;
    const { isLoading, loadingProgress, loadingMessage, currentItem, totalItems, isComplete, error, showDialog, isThrottled } = this.state;

    const dialogContentProps = {
      type: DialogType.normal,
      title: isComplete ? 'Template Loaded Successfully' : 'Initialize Project Folder',
      subText: isComplete
        ? `All template files and folders have been successfully copied to ${projectName}.`
        : locationName
          ? `Initialize "${projectName}" in "${locationName}" with the standard project template?`
          : `Initialize "${projectName}" with the standard project template?`
    };

    return (
      <Dialog
        hidden={!showDialog}
        onDismiss={this._handleCancel}
        dialogContentProps={dialogContentProps}
        modalProps={{
          isBlocking: isLoading,
          styles: { main: { maxWidth: 550 } }
        }}
      >
        <DialogContent>
          {!isComplete && !error && (
            <Stack tokens={stackTokens}>
              {!isLoading && (
                <>
                  <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                    <Icon iconName="Info" styles={{ root: { fontSize: 20, color: '#0078d4' } }} />
                    <Text>
                      This will copy all files and folders from the Templates library into this project folder.
                    </Text>
                  </Stack>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Note: This action cannot be undone automatically.
                  </Text>
                </>
              )}

              {isLoading && (
                <Stack tokens={stackTokens}>
                  {isThrottled && (
                    <Stack
                      styles={{
                        root: {
                          backgroundColor: '#fff4ce',
                          padding: '12px',
                          borderRadius: '4px',
                          border: '1px solid #ffc83d'
                        }
                      }}
                      tokens={{ childrenGap: 10 }}
                    >
                      <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                        <Icon
                          iconName="Warning"
                          styles={{ root: { fontSize: 24, color: '#d83b01' } }}
                        />
                        <Stack>
                          <Text variant="medium" styles={{ root: { fontWeight: FontWeights.semibold, color: '#d83b01' } }}>
                            SharePoint Throttling Detected
                          </Text>
                          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                            {loadingMessage}
                          </Text>
                        </Stack>
                      </Stack>
                    </Stack>
                  )}

                  <ProgressIndicator
                    label="Loading Template"
                    description={!isThrottled ? loadingMessage : ''}
                    percentComplete={loadingProgress / 100}
                    progressHidden={false}
                  />
                  <Stack horizontal tokens={{ childrenGap: 20 }} horizontalAlign="space-between">
                    <Text variant="medium" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                      {loadingProgress}% Complete
                    </Text>
                    {totalItems > 0 && (
                      <Text variant="medium" styles={{ root: { fontWeight: FontWeights.semibold, color: '#0078d4' } }}>
                        Item {currentItem} of {totalItems}
                      </Text>
                    )}
                  </Stack>
                  <Text variant="small" styles={{ root: { color: '#605e5c', fontStyle: 'italic' } }}>
                    Please do not close this tab while the template is loading...
                  </Text>
                </Stack>
              )}
            </Stack>
          )}

          {isComplete && (
            <Stack tokens={stackTokens} horizontalAlign="center">
              <Icon
                iconName="CheckMark"
                styles={{
                  root: {
                    fontSize: 48,
                    color: '#107c10',
                    border: '3px solid #107c10',
                    borderRadius: '50%',
                    padding: 10
                  }
                }}
              />
              <Text variant="large" styles={{ root: { color: '#107c10' } }}>
                All files and folders have been copied successfully!
              </Text>
            </Stack>
          )}

          {error && (
            <Stack tokens={stackTokens}>
              <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center" horizontalAlign="space-between">
                <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                  <Icon iconName="Error" styles={{ root: { fontSize: 20, color: '#a4262c' } }} />
                  <Text styles={{ root: { color: '#a4262c', fontWeight: FontWeights.semibold } }}>
                    Error loading template:
                  </Text>
                </Stack>
                <DefaultButton
                  text="Copy Error"
                  iconProps={{ iconName: 'Copy' }}
                  onClick={() => this._copyErrorToClipboard(error)}
                  styles={{ root: { minWidth: 'auto' } }}
                />
              </Stack>
              <div style={{
                backgroundColor: '#fde7e9',
                border: '1px solid #a4262c',
                borderRadius: '4px',
                padding: '12px',
                maxHeight: '300px',
                overflowY: 'auto',
                userSelect: 'text'
              }}>
                <Text
                  variant="small"
                  styles={{
                    root: {
                      color: '#a4262c',
                      whiteSpace: 'pre-wrap',
                      fontFamily: 'Consolas, Monaco, "Courier New", monospace',
                      userSelect: 'text'
                    }
                  }}
                >
                  {error}
                </Text>
              </div>
            </Stack>
          )}
        </DialogContent>

        <DialogFooter>
          {!isLoading && !isComplete && !error && (
            <>
              <PrimaryButton onClick={this._handleLoadTemplate} text="Load Template" style={{ marginRight: '30px' }} />
              <DefaultButton onClick={this._handleCancel} text="Cancel" />
            </>
          )}
          {(isComplete || error) && (
            <PrimaryButton onClick={this._handleClose} text="Close" />
          )}
        </DialogFooter>
      </Dialog>
    );
  }

  private _handleLoadTemplate = async (): Promise<void> => {
    this.setState({ isLoading: true, loadingProgress: 0, loadingMessage: 'Initializing...', currentItem: 0, totalItems: 0, isThrottled: false });

    try {
      await this.props.loadTemplate(
        this.props.folderUrl,
        (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => {
          this.setState({
            loadingProgress: percent,
            loadingMessage: message,
            currentItem: currentItem || 0,
            totalItems: totalItems || 0,
            isThrottled: isThrottled || false
          });
        }
      );

      this.setState({
        isLoading: false,
        isComplete: true,
        loadingProgress: 100,
        loadingMessage: 'Complete!'
      });

      // Auto-close after 3 seconds
      setTimeout(() => {
        this._handleClose();
      }, 3000);

    } catch (error) {
      console.error('[TemplateLoaderDialog] Error loading template:', error);

      // Extract detailed error message
      let errorMessage = 'An unexpected error occurred';
      if (error.message) {
        errorMessage = error.message;
      } else if (typeof error === 'string') {
        errorMessage = error;
      }

      this.setState({
        isLoading: false,
        error: errorMessage
      });
    }
  }

  private _handleCancel = (): void => {
    if (!this.state.isLoading) {
      this.props.onCancel();
      this.setState({ showDialog: false });
      this.props.onClose();
    }
  }

  private _handleClose = (): void => {
    this.setState({ showDialog: false });
    this.props.onClose();
  }

  private _copyErrorToClipboard = (text: string): void => {
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text).then(
        () => {
          console.log('[TemplateLoaderDialog] Error text copied to clipboard');
        },
        (err) => {
          console.error('[TemplateLoaderDialog] Could not copy text: ', err);
        }
      );
    } else {
      // Fallback for older browsers
      const textArea = document.createElement('textarea');
      textArea.value = text;
      textArea.style.position = 'fixed';
      textArea.style.opacity = '0';
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      try {
        document.execCommand('copy');
        console.log('[TemplateLoaderDialog] Error text copied to clipboard (fallback)');
      } catch (err) {
        console.error('[TemplateLoaderDialog] Fallback copy failed: ', err);
      }
      document.body.removeChild(textArea);
    }
  }
}

export default class TemplateLoaderDialog extends BaseDialog {
  private _folderUrl: string;
  private _programName: string;
  private _locationName: string;
  private _projectName: string;
  private _loadTemplateCallback: (folderUrl: string, onProgress?: (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => void) => Promise<void>;
  private _onCancelCallback: () => void;
  private _autoStart: boolean;

  constructor(
    folderUrl: string,
    programName: string,
    locationName: string,
    projectName: string,
    loadTemplateCallback: (folderUrl: string, onProgress?: (percent: number, message: string, currentItem?: number, totalItems?: number, isThrottled?: boolean) => void) => Promise<void>,
    onCancelCallback: () => void,
    autoStart: boolean = false
  ) {
    super();
    this._folderUrl = folderUrl;
    this._programName = programName;
    this._locationName = locationName;
    this._projectName = projectName;
    this._loadTemplateCallback = loadTemplateCallback;
    this._onCancelCallback = onCancelCallback;
    this._autoStart = autoStart;
  }

  public render(): void {
    ReactDOM.render(
      <TemplateLoaderDialogContent
        folderUrl={this._folderUrl}
        programName={this._programName}
        locationName={this._locationName}
        projectName={this._projectName}
        loadTemplate={this._loadTemplateCallback}
        onCancel={this._onCancelCallback}
        onClose={this.close.bind(this)}
        autoStart={this._autoStart}
      />,
      this.domElement
    );
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}