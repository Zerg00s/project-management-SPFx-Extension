declare interface IProjectManagerCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ProjectManagerCommandSetStrings' {
  const strings: IProjectManagerCommandSetStrings;
  export = strings;
}
