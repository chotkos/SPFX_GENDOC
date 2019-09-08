declare interface IGenDocCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'GenDocCommandSetStrings' {
  const strings: IGenDocCommandSetStrings;
  export = strings;
}
