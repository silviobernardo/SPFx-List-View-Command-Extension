declare interface ISpfxListViewCommandExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SpfxListViewCommandExtensionCommandSetStrings' {
  const strings: ISpfxListViewCommandExtensionCommandSetStrings;
  export = strings;
}
