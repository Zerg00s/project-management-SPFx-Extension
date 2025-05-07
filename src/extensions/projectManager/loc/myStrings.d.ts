declare interface IProjectManagerCommandSetStrings {
  Command1: string;
  Command2: string;
  CreateProject: string;
  SelectTemplate: string;
  EnterProjectName: string;
  Create: string;
  Cancel: string;
  CreatingProject: string;
  ProjectCreated: string;
  ErrorCreatingProject: string;
}

declare module 'ProjectManagerCommandSetStrings' {
  const strings: IProjectManagerCommandSetStrings;
  export = strings;
}
