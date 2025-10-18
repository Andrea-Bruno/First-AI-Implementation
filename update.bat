@echo off
REM Git synchronization script: initializes, pulls, commits, and pushes code to remote repository

REM Set the remote repository URL
set REMOTE_URL=https://github.com/Andrea-Bruno/First-AI-Implementation.git

REM Check if the .git folder exists (i.e., if the repo is already initialized)
IF NOT EXIST ".git" (
    echo Initializing Git repository...
    git init
    git remote add origin %REMOTE_URL%
    git branch -M main
)

REM Pull latest changes from remote to avoid conflicts
echo Pulling latest changes from remote repository...
git pull origin main --allow-unrelated-histories

REM Add all modified and new files to staging
echo Adding files to staging...
git add .

REM Commit changes with a default message
echo Committing changes...
git commit -m "Sync update"

REM Push local changes to the remote repository
echo Pushing changes to GitHub...
git push origin main

echo Synchronization complete!
pause
