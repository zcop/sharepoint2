sharepoint backend for Nextcloud  
Base on Oauth2 + Graph API  
Currently works on Nextcloud 32(Hub 25)

Cache operations:
- Auto warmup queue after SharePoint2 mount create/update (root scan + level-1 split scan).
- Manual warmup: `sudo -u www-data php occ sharepoint2:cache:warmup <mount_id>`
- Manual single path scan: `sudo -u www-data php occ sharepoint2:cache:warmup <mount_id> --path="FolderA"`
- Manual full scan: `sudo -u www-data php occ sharepoint2:cache:warmup <mount_id> --full`
