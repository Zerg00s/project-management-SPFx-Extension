import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { 
    Panel, 
    PanelType, 
    TextField, 
    PrimaryButton, 
    DefaultButton,
    Stack,      
    MessageBar, 
    MessageBarType,
    Spinner,
    SpinnerSize,
    IStackTokens,
    Text,
    FontWeights  
} from 'office-ui-fabric-react';
import { IProjectTemplate } from './ProjectManagerCommandSet';

export interface IProjectTemplatePanelProps {
    templates: IProjectTemplate[];
    createProject: (template: IProjectTemplate, projectName: string) => Promise<void>;
    closePanel: () => void;
}

interface IProjectTemplatePanelState {
    selectedTemplate: IProjectTemplate | null;
    projectName: string;
    error: string;
    isCreating: boolean;
    isCreated: boolean;
}

// Define spacing for stack items
const stackTokens: IStackTokens = { childrenGap: 15 };

// Define the Template Panel component
class TemplatePanelContent extends React.Component<IProjectTemplatePanelProps, IProjectTemplatePanelState> {
    constructor(props: IProjectTemplatePanelProps) {
        super(props);
        this.state = {
            selectedTemplate: null,
            projectName: '',
            error: '',
            isCreating: false,
            isCreated: false
        };
    }

    public render(): JSX.Element {
        const { templates } = this.props;
        const { selectedTemplate, projectName, error, isCreating, isCreated } = this.state;

        return (
            <Panel
                isOpen={true}
                type={PanelType.medium}
                headerText="Create a Project"
                closeButtonAriaLabel="Close"
                onDismiss={this.props.closePanel}
                isLightDismiss={!isCreating}
            >
                {isCreated ? (
                    <Stack tokens={stackTokens}>
                        <MessageBar messageBarType={MessageBarType.success}>
                            Project "{projectName}" created successfully!
                        </MessageBar>
                        <PrimaryButton onClick={this.props.closePanel} text="Close" />
                    </Stack>
                ) : (
                    <Stack tokens={stackTokens}>
                        {error && (
                            <MessageBar messageBarType={MessageBarType.error}>
                                {error}
                            </MessageBar>
                        )}

                        {isCreating ? (
                            <Stack tokens={stackTokens} horizontalAlign="center">
                                <Spinner size={SpinnerSize.large} label={`Creating project "${projectName}"...`} />
                            </Stack>
                        ) : (
                            <>
                                <Text styles={{ root: { fontWeight: FontWeights.semibold } }}>
                                    Step 1: Select a Project Template
                                </Text>
                                
                                <Stack horizontal wrap tokens={stackTokens}>
                                    {templates.map((template) => (
                                        <Stack.Item key={template.FileRef} styles={{ root: { width: '150px', marginBottom: '15px' } }}>
                                            <div 
                                                style={{ 
                                                    border: selectedTemplate && selectedTemplate.FileRef === template.FileRef ? '2px solid #0078d4' : '1px solid #c8c8c8',
                                                    padding: '15px', 
                                                    cursor: 'pointer',
                                                    backgroundColor: selectedTemplate && selectedTemplate.FileRef === template.FileRef ? '#f0f8ff' : 'white',
                                                    textAlign: 'center',
                                                    height: '100px',
                                                    display: 'flex',
                                                    flexDirection: 'column',
                                                    justifyContent: 'center',
                                                    borderRadius: '2px'
                                                }}
                                                onClick={() => this.selectTemplate(template)}
                                            >
                                                <i className="ms-Icon ms-Icon--FolderHorizontal" style={{ fontSize: '32px', color: '#0078d4', marginBottom: '8px' }}></i>
                                                <Text>{template.FileLeafRef}</Text>
                                            </div>
                                        </Stack.Item>
                                    ))}
                                </Stack>
                                
                                <Text styles={{ root: { fontWeight: FontWeights.semibold } }}>
                                    Step 2: Enter Project Name
                                </Text>
                                
                                <TextField 
                                    label="Project Name" 
                                    value={projectName}
                                    onChange={this.onProjectNameChange}
                                    disabled={!selectedTemplate}
                                    placeholder="Enter a name for your new project"
                                    required
                                />
                                
                                <Stack horizontal tokens={{ childrenGap: 10 }}>
                                    <PrimaryButton 
                                        text="Create Project" 
                                        onClick={this.createProject}
                                        disabled={!selectedTemplate || !projectName}
                                    />
                                    <DefaultButton 
                                        text="Cancel" 
                                        onClick={this.props.closePanel} 
                                    />
                                </Stack>
                            </>
                        )}
                    </Stack>
                )}
            </Panel>
        );
    }

    private selectTemplate = (template: IProjectTemplate): void => {
        this.setState({ selectedTemplate: template });
    }

    private onProjectNameChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ projectName: newValue || '' });
    }

    private createProject = async (): Promise<void> => {
        const { selectedTemplate, projectName } = this.state;
        
        if (!selectedTemplate || !projectName.trim()) {
            this.setState({ error: 'Please select a template and provide a project name.' });
            return;
        }

        this.setState({ isCreating: true, error: '' });

        try {
            await this.props.createProject(selectedTemplate, projectName);
            this.setState({ isCreated: true, isCreating: false });
        } catch (error) {
            this.setState({ 
                error: `Error creating project: ${error.message}`, 
                isCreating: false 
            });
        }
    }
}

// Define the Panel Dialog class that extends BaseDialog
export default class ProjectTemplatePanel extends BaseDialog {
    private _templates: IProjectTemplate[];
    private _createProjectCallback: (template: IProjectTemplate, projectName: string) => Promise<void>;

    constructor(templates: IProjectTemplate[], createProjectCallback: (template: IProjectTemplate, projectName: string) => Promise<void>) {
        super();
        this._templates = templates;
        this._createProjectCallback = createProjectCallback;
    }

    public render(): void {
        ReactDOM.render(
            <TemplatePanelContent 
                templates={this._templates} 
                createProject={this._createProjectCallback}
                closePanel={this.close.bind(this)}
            />, 
            this.domElement
        );
    }

    protected onAfterClose(): void {
        super.onAfterClose();
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}
