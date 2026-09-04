# Public deployment notes (v18)

## Vercel

The project includes `vercel.json` and can be deployed from the project root with:

```bash
vercel --prod
```

or via Git integration after pushing this folder to a repository.

## Translation

v18 uses online translation only through `deep-translator`'s Google Translate integration. The application does not require an API key or paid Cloud Translation credentials. This is not an official paid Google Cloud Translation API integration and availability/rate limits are controlled by the upstream service. Text submitted for translation is sent to the external translation service.

## Storage on serverless hosting

The app separates browser workspaces with an HttpOnly cookie, but Vercel's writable temporary filesystem is ephemeral. Use **Save Project** (`.paperproj.zip`) for durable work and re-open the saved project when returning later. For durable cloud persistence, attach an external database/blob store in a future version.
