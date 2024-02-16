## Sample goal

This sample will add and manage custom Commands into a SharePoint List and allow you to archive some records (one by one).


## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)


## HOW TO RUN

1. Create two lists on SharePoint with two columns: Title (text) and Age (number)
2. Add some data on first list
3. Open repository folder on Visual Studio Code
4. Go to "config/serve.json" file and update ...
    2.1. "pageUrl" value to the first SharePoint List URL (the one that you have added some data)
    2.1. "properties -> archiveList" with the title of the second list that you have just created and it's already empty
5. Open a terminal
6. Run "npm install"
7. Run "gulp trust-dev-cert"
8. Run "gulp serve" command.
9. Test "Command One" button and "Archive" button (it will be available once you have selected a single row)
