import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import ProjectTemplatePanel from './ProjectTemplatePanel';

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

export interface IProjectTemplate {
  Title: string;
  FileRef: string;
  FileLeafRef: string;
}

export default class ProjectManagerCommandSet extends BaseListViewCommandSet<IProjectManagerCommandSetProperties> {
  private _sp: SPFI;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ProjectManagerCommandSet');
    
    // Initialize PnP
    this._sp = spfi().using(SPFx(this.context));
    
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    const createProjectCommand: Command = this.tryGetCommand('CREATE_PROJECT');
    if (createProjectCommand) {
      // Filter to only show on Projects library
      let isProjectsLibrary = false;
      
      // Safely check if list info exists
      if (this.context.pageContext && this.context.pageContext.list) {
        const listUrl = this.context.pageContext.list.serverRelativeUrl || '';
        const listTitle = this.context.pageContext.list.title || '';
        
        isProjectsLibrary = 
          (listTitle === 'Projects') || 
          (listUrl.indexOf('/Projects') > -1 && listUrl.indexOf('/Projects') === listUrl.length - '/Projects'.length) || 
          (listUrl.indexOf('/sites/ProjectsRepository/Projects') > -1);
      }
      
      createProjectCommand.visible = isProjectsLibrary;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'CREATE_PROJECT':
        this._showProjectTemplatePanel();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _showProjectTemplatePanel(): void {
    // Fetch templates and show panel
    this._fetchProjectTemplates()
      .then((templates: IProjectTemplate[]) => {
        if (templates.length === 0) {
          // eslint-disable-next-line no-void
          void Dialog.alert('No project templates found. Please create templates in the ProjectTemplates library.');
          return;
        }

        // Show the template panel
        const panel = new ProjectTemplatePanel(templates, this._createProjectPnP.bind(this));
        // eslint-disable-next-line no-void
        void panel.show();
      })
      .catch((error: Error) => {
        // eslint-disable-next-line no-void
        void Dialog.alert(`Error fetching project templates: ${error.message}`);
      });
  }

  private _fetchProjectTemplates(): Promise<IProjectTemplate[]> {
    const endpoint = "/sites/ProjectsRepository/_api/web/GetFolderByServerRelativeUrl('ProjectTemplates')/Folders";
  
    return this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => {
      if (!response.ok) {
        throw new Error(`HTTP error: ${response.status} ${response.statusText}`);
      }
      return response.json();
    })
    .then((data: any) => {
      // Map response to IProjectTemplate[]
      // exclude Folders with "Forms" in the name
      data.value = data.value.filter((item: any) => !item.Name.includes("Forms"));
      // sort by name
      data.value.sort((a: any, b: any) => a.Name.localeCompare(b.Name));
      return data.value.map((item: any) => ({
        Title: item.Name,
        FileRef: item.ServerRelativeUrl,
        FileLeafRef: item.Name
      })) as IProjectTemplate[];
    });
  }
  

  private async _createProjectPnP(template: IProjectTemplate, projectName: string): Promise<void> {
    try {
      // Get the correct server relative URL paths
      const sourceUrl = template.FileRef;  // This is already server-relative
      const destinationUrl = `/sites/ProjectsRepository/Projects/${projectName}`;
      
      console.log('Source path:', sourceUrl);
      console.log('Destination path:', destinationUrl);

      // Use PnP.js to copy the folder
      await this._sp.web.getFolderByServerRelativePath(sourceUrl).copyByPath(destinationUrl, true);
      
      return Promise.resolve();
    } catch (error) {
      console.error('Project creation error:', error);
      throw error;
    }
  }
}