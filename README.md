# RichOffice
Discord Rich Presence for Microsoft Office

## Binary
- Depending on which application you want to use, individual Discord Developer Applications required to be created.
- Each application requires an OAuth2 URL of **http://127.0.0.1**.
- Each application also requires an art asset (icon of the app) named using lowercase convention: "excel" for example.
- Download and extract to a single folder.
- Rename the given default config file to **config.json**, fill required fields with OAuth2 Client IDs and move to same directory as binary.
- Open executable.

## Building
- Download and extract files to a folder.
- Script depends on:
>**DiscordRichPresence** by Lachee at https://github.com/Lachee/discord-rpc-csharp. <br/>
>**Newtonsoft.Json** by James Newton-King at https://www.newtonsoft.com/json.
- Follow steps from the *Binary* section.
