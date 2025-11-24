# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PlcTestSuite_Runner is a .NET Framework 4.7.2 console application that automates TwinCAT XAE (eXtended Automation Engineering) projects. It programmatically opens TwinCAT solutions, activates configurations, and starts the TwinCAT runtime for PLC test automation.

## Architecture

The application consists of three main components:

1. **Program.Main** (Program.cs:15-51): Entry point that handles command-line arguments and error reporting
   - Expects two arguments: solution file path (relative to parent directory) and project name
   - Implements structured error handling with COM and general exception handling

2. **Automation Class** (Program.cs:53-148): Core automation logic
   - `ActivateProject(string project, string projectName)`: Opens a TwinCAT solution via DTE (Development Tools Environment) COM interface
   - Searches for the specified project name within the solution
   - Uses ITcSysManager3 to activate configuration and start TwinCAT runtime
   - Hard-coded wait times: 50 seconds for project load, 10 seconds for TwinCAT startup

3. **MessageFilter Class** (Program.cs:150-236): COM interop thread handling
   - Implements IOleMessageFilter to handle COM threading issues
   - Automatically retries rejected COM calls (up to 99 times immediately)
   - Registered at application start, revoked on cleanup

## Key Dependencies

- **EnvDTE**: Visual Studio automation COM library for manipulating solutions/projects
- **TCatSysManagerLib**: TwinCAT System Manager COM library (version 3.4) for PLC operations
  - GUID: {3C49D6C3-93DC-11D0-B162-00A0248C244B}

## Building and Running

Build the project:
```
msbuild PlcTestSuite_Runner\PlcTestSuite_Runner.csproj /p:Configuration=Debug
msbuild PlcTestSuite_Runner\PlcTestSuite_Runner.csproj /p:Configuration=Release
```

Run the application:
```
cd PlcTestSuite_Runner\bin\Debug
PlcTestSuite_Runner.exe <solution-file-name.sln> <project-name>
```

The solution file path is resolved relative to the parent directory of the call directory.

## Important Implementation Details

- **[STAThread]** attribute is required on Main for COM interop
- The application expects TwinCAT XAE Shell 17.0 (ProgID: "TcXaeShell.DTE.17.0")
- DTE instance is created with visible UI (`dte.SuppressUI = false`, `dte.MainWindow.Visible = true`)
- Project name must be specified as a command-line argument
- COM interop types are embedded in the assembly (EmbedInteropTypes=True)

## Common Error Scenarios

- COMException: TwinCAT not installed, version mismatch, or licensing issues
- FileNotFoundException: Invalid solution path
- InvalidOperationException: Specified project not found or ITcSysManager3 interface unavailable
- Exit code 1 indicates any error condition
