declare interface ISpFxListViewCommandExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SpFxListViewCommandExtensionCommandSetStrings' {
  const strings: ISpFxListViewCommandExtensionCommandSetStrings;
  export = strings;
}
