# Attendance Portal Enterprise Updater V15

## Important: preserve existing business data
Do not delete or overwrite existing JSON, XLSX, or `uploads/` files. Copy the new `app.py` into the same portal folder so the application continues using the existing users, leave history, approved overtime, attendance corrections, salary records, and uploaded documents.

Before deployment, make a complete copy of the current portal folder.

## Package contents
- `app.py`: updated portal
- `requirements.txt`: required Python packages
- `.env.example`: environment-variable checklist, with no real secrets
- `install_windows.bat` and `start_portal.bat`
- `install_linux.sh` and `start_linux.sh`
- `verify_install.py`: offline syntax, route, and folder check

## Windows installation
1. Install Python 3.11 or 3.12 and enable **Add Python to PATH**.
2. Extract this ZIP into the existing portal folder. Keep existing data files and `uploads/`.
3. Run `install_windows.bat` once.
4. Configure secrets using Windows Environment Variables. Do not put real tokens inside `app.py`.
5. Restart the terminal or service, then run `start_portal.bat`.
6. Open `/api/build` and verify build `2026-10-12-ENTERPRISE-PORTAL-UPDATER-V15`.
7. Login as developer and open `/developer/portal_update`.

Example persistent variables in Administrator CMD:

```bat
setx SECRET_KEY "CHANGE-TO-A-LONG-RANDOM-SECRET"
setx GITHUB_TOKEN "YOUR-FINE-GRAINED-GITHUB-TOKEN"
setx GITHUB_REPO_NAME "devilsuza/Gamekfme"
setx GITHUB_BRANCH "main"
setx PORTAL_UPDATE_PASSWORD "STRONG-DEPLOYMENT-PASSWORD"
setx PORTAL_UPDATE_APPROVAL_PASSWORD "SECOND-APPROVAL-PASSWORD"
setx PORTAL_BACKUP_RETENTION "10"
setx PORTAL_UPDATE_MAX_MB "8"
```

Close and reopen CMD after `setx`.

## GitHub token permissions
Use a fine-grained token limited to the required repository. Repository permission **Contents: Read and write** is needed for document uploads. Keep the token in an environment variable or hosting secret.

## Existing data that must remain beside app.py
Depending on the deployed portal, preserve all files already present, especially:
- users database JSON
- `leave_requests.json`
- `leave_records.xlsx`
- overtime approvals JSON
- salary slips JSON
- attendance corrections and overrides
- roster, machines, manual and duplicate punch data
- entire `uploads/` directory

The deployment center creates `portal_updates/` and `portal_backups/` automatically.

## Validation
Run:

```bat
python verify_install.py
```

## Production recommendation
For reliable automatic restart, run the portal under NSSM/Windows Service, systemd, Docker, or another process supervisor. The built-in restart works when the Python process is allowed to replace itself, but a service supervisor is safer for production.
