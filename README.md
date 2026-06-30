# AttendX

AttendX is a desktop classroom management app for tracking attendance, marks, weighted percentages, grades, and student risk insights. It is built with React, Vite, Electron, and XLSX export/import support.

The app supports two class types:

- **Attendance classes**: mark students present/absent by date, track attendance percentage, and identify students below required attendance.
- **Marks classes**: create assessments, assign weights, enter marks, calculate weighted percentage, assign grades, and identify students below the weighted-score threshold.

## Features

- Create separate attendance and marks classes.
- Import student lists from Excel or CSV.
- Export individual classes or all classes to Excel.
- Generate attendance defaulter lists.
- Download and upload marks templates.
- Always-visible weighted percentage column for marks classes.
- Grade calculation based on weighted score.
- Show Insights panel with maximize/minimize support.
- Export Insights for attendance and marks risk lists.
- Undo and redo support.
- Dark mode.
- Local desktop data storage through Electron.

## Download From GitHub

If the repository has a prepared installer attached to a GitHub Release:

1. Open the GitHub repository.
2. Go to **Releases**.
3. Download `AttendX Setup 1.0.0.exe`.
4. Run the installer.

If there is no release installer yet, build it from source using the steps below.

## Requirements

- Windows 10 or later
- Node.js 18 or later
- npm

Check your versions:

```powershell
node -v
npm -v
```

## Install Dependencies

Clone or download the repository, then open PowerShell in the project folder:

```powershell
cd AttendX
npm.cmd install
```

On Windows, `npm.cmd` is recommended because PowerShell may block `npm.ps1` depending on execution policy.

## Run In Development Mode

For browser-based development:

```powershell
npm.cmd run dev
```

Then open the local URL shown by Vite, usually:

```text
http://localhost:5173
```

For Electron development:

```powershell
npm.cmd run electron:dev
```

## Create The Windows Installer

Build the production app and create the Windows installer:

```powershell
npm.cmd run electron:build
```

After a successful build, the installer will be created here:

```text
dist-electron\AttendX Setup 1.0.0.exe
```

The unpacked app is also generated here:

```text
dist-electron\win-unpacked
```

## Install AttendX

Run the installer:

```powershell
& ".\dist-electron\AttendX Setup 1.0.0.exe"
```

For silent installation:

```powershell
& ".\dist-electron\AttendX Setup 1.0.0.exe" /S
```

After installation, AttendX is usually installed at:

```text
%LOCALAPPDATA%\Programs\AttendX\AttendX.exe
```

You can also launch it from the Windows Start Menu.

## Data Safety

AttendX stores app data locally on your computer. Building or reinstalling the app does not intentionally delete your classes, attendance, marks, weights, or grades.

Still, before major updates, it is a good habit to export your classes from inside AttendX:

```text
Export All Classes
```

## Build Scripts

| Command | Purpose |
| --- | --- |
| `npm.cmd run dev` | Start the Vite development server |
| `npm.cmd run build` | Build the React app |
| `npm.cmd run electron:dev` | Run the app in Electron development mode |
| `npm.cmd run electron:build` | Build the Windows installer |

## Project Structure

```text
AttendX
|-- src
|   |-- App.jsx
|   |-- App.css
|   `-- main.jsx
|-- public
|-- electron.cjs
|-- preload.cjs
|-- package.json
`-- dist-electron
```

## Tech Stack

- React
- Vite
- Electron
- electron-builder
- XLSX

## License

Add your license here before publishing the repository publicly.
