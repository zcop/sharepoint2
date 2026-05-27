# SharePoint2 for Nextcloud

`sharepoint2` is an external storage backend for Nextcloud that mounts Microsoft SharePoint libraries using OAuth 2.0 (Microsoft Entra ID) and Microsoft Graph API.

## Highlights

- OAuth 2.0 app-only authentication (tenant ID, client ID, client secret)
- SharePoint library mount through Nextcloud External Storage
- Read-only storage behavior for safe integration
- Access token lifecycle handling with refresh support
- Cache warmup workflow for faster folder navigation on large trees
- Operational commands for targeted or full warmup scans

## Compatibility

- Designed for modern Nextcloud versions using `files_external`
- Validated in this project workflow on Nextcloud Hub 8/9 environments

## Requirements

- Nextcloud with `files_external` enabled
- Microsoft Entra ID app registration with Graph permissions suitable for SharePoint library read access
- Tenant ID, Client ID, and Client Secret
- Network access from Nextcloud server to `graph.microsoft.com`

## Installation

1. Copy this app folder to your Nextcloud `apps/` directory as `sharepoint2`.
2. Enable the app:

```bash
sudo -u www-data php occ app:enable sharepoint2
```

3. In Nextcloud Admin settings, go to External storage and create a new mount using `SharePoint2`.

## Configuration Fields

- `Site URL`: SharePoint site URL (example: `https://tenant.sharepoint.com/sites/MySite`)
- `Library`: library name or library-relative path (example: `Documents` or `Documents/SubFolder`)
- `Tenant`: Entra tenant ID (GUID)
- `Client ID`: Entra application (client) ID
- `Client Secret`: Entra application secret

## Cache Warmup Operations

Warmup helps prebuild file cache state and improve browsing performance.

- Manual warmup (default root strategy):

```bash
sudo -u www-data php occ sharepoint2:cache:warmup <mount_id>
```

- Warmup a single path:

```bash
sudo -u www-data php occ sharepoint2:cache:warmup <mount_id> --path="FolderA"
```

- Full warmup:

```bash
sudo -u www-data php occ sharepoint2:cache:warmup <mount_id> --full
```

Notes:
- Mount create/update can enqueue warmup automatically.
- For very large libraries, schedule warmup during low-traffic windows.

## Operational Notes

- Storage is intended to be read-only from Nextcloud.
- If a file appears in directory listing but cannot be opened, verify Entra app permissions, site/library path correctness, and Graph API reachability (including firewall rules).

## Troubleshooting

- Check Nextcloud logs:

```bash
sudo tail -f /var/www/nextcloud/data/nextcloud.log
```

- Validate app is enabled:

```bash
sudo -u www-data php occ app:list | grep sharepoint2
```

- List external mounts:

```bash
sudo -u www-data php occ files_external:list
```

## License

See repository license terms.
