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
  loadTemplate: (folderUrl: string, onProgress?: (percent: number, message: string) => void) => Promise<void>;
  onCancel: () => void;
  onClose: () => void;
}

interface ITemplateLoaderDialogState {
  isLoading: boolean;
  loadingProgress: number;
  loadingMessage: string;
  isComplete: boolean;
  error: string;
  showDialog: boolean;
}

const stackTokens: IStackTokens = { childrenGap: 20 };

class TemplateLoaderDialogContent extends React.Component<ITemplateLoaderDialogProps, ITemplateLoaderDialogState> {
  constructor(props: ITemplateLoaderDialogProps) {
    super(props);
    this.state = {
      isLoading: false,
      loadingProgress: 0,
      loadingMessage: '',
      isComplete: false,
      error: '',
      showDialog: true
    };
  }

  public render(): JSX.Element {
    const { programName, locationName, projectName } = this.props;
    const { isLoading, loadingProgress, loadingMessage, isComplete, error, showDialog } = this.state;

    const dialogContentProps = {
      type: DialogType.normal,
      title: isComplete ? 'Template Loaded Successfully' : 'Empty Project Folder Detected',
      subText: isComplete
        ? `The template has been successfully loaded into ${projectName}.`
        : `The folder "${projectName}" in location "${locationName}" (${programName} program) is empty. Would you like to load the project template?`
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
                  <ProgressIndicator
                    label="Loading Template"
                    description={loadingMessage}
                    percentComplete={loadingProgress / 100}
                    progressHidden={false}
                  />
                  <Text variant="medium" styles={{ root: { fontWeight: FontWeights.semibold } }}>
                    {loadingProgress}% Complete
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
              <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                <Icon iconName="Error" styles={{ root: { fontSize: 20, color: '#a4262c' } }} />
                <Text styles={{ root: { color: '#a4262c', fontWeight: FontWeights.semibold } }}>
                  Error loading template:
                </Text>
              </Stack>
              <div style={{
                backgroundColor: '#fde7e9',
                border: '1px solid #a4262c',
                borderRadius: '4px',
                padding: '12px',
                maxHeight: '300px',
                overflowY: 'auto'
              }}>
                <Text
                  variant="small"
                  styles={{
                    root: {
                      color: '#a4262c',
                      whiteSpace: 'pre-wrap',
                      fontFamily: 'Consolas, Monaco, "Courier New", monospace'
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
    this.setState({ isLoading: true, loadingProgress: 0, loadingMessage: 'Initializing...' });

    try {
      await this.props.loadTemplate(
        this.props.folderUrl,
        (percent: number, message: string) => {
          this.setState({
            loadingProgress: percent,
            loadingMessage: message
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
}

export default class TemplateLoaderDialog extends BaseDialog {
  private _folderUrl: string;
  private _programName: string;
  private _locationName: string;
  private _projectName: string;
  private _loadTemplateCallback: (folderUrl: string, onProgress?: (percent: number, message: string) => void) => Promise<void>;
  private _onCancelCallback: () => void;

  constructor(
    folderUrl: string,
    programName: string,
    locationName: string,
    projectName: string,
    loadTemplateCallback: (folderUrl: string, onProgress?: (percent: number, message: string) => void) => Promise<void>,
    onCancelCallback: () => void
  ) {
    super();
    this._folderUrl = folderUrl;
    this._programName = programName;
    this._locationName = locationName;
    this._projectName = projectName;
    this._loadTemplateCallback = loadTemplateCallback;
    this._onCancelCallback = onCancelCallback;
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
      />,
      this.domElement
    );
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}