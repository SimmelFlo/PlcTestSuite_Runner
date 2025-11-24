# PlcTestSuite Runner

Automates TwinCAT XAE projects by opening solutions, activating configurations, and starting the TwinCAT runtime.

## Prerequisites

- TwinCAT XAE Shell 17.0
- .NET Framework 4.7.2+

## Usage

```bash
PlcTestSuite_Runner.exe <solution-file.sln> <project-name>
```

**Important**: Run the executable from a subdirectory. The application looks for the solution file in the parent directory.

### Example

```
MyProject/
├── MyTwinCATSolution.sln
└── TestRunner/
    └── PlcTestSuite_Runner.exe
```

```bash
cd MyProject\TestRunner
PlcTestSuite_Runner.exe MyTwinCATSolution.sln CoreProject
```

## What It Does

1. Opens TwinCAT XAE Shell
2. Loads your solution (waits 50 seconds)
3. Finds the specified project by name
4. Activates configuration and starts TwinCAT runtime (waits 10 seconds)

## Building

```bash
# Debug
msbuild PlcTestSuite_Runner\PlcTestSuite_Runner.csproj /p:Configuration=Debug

# Release
msbuild PlcTestSuite_Runner\PlcTestSuite_Runner.csproj /p:Configuration=Release
```

Output: `PlcTestSuite_Runner\bin\[Debug|Release]\PlcTestSuite_Runner.exe`

## Troubleshooting

| Error | Solution |
|-------|----------|
| "Usage: PlcTestSuite_Runner.exe..." | Provide both solution file and project name as arguments |
| "COM Exception occurred" | Verify TwinCAT XAE Shell 17.0 is installed |
| "Project file not found" | Check solution file exists in parent directory |
| "[ProjectName] not found" | Ensure solution contains the specified project name |

## License

MIT License - See [LICENSE](LICENSE) file for details.
