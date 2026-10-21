# Publishing OpenDocx to NuGet

This document explains how to publish the OpenDocx package to NuGet.org.

## Prerequisites

1. **NuGet Account**: You need a NuGet.org account. Sign up at https://www.nuget.org if you don't have one.

2. **API Key**: Create an API key from your NuGet account:
   - Go to https://www.nuget.org
   - Sign in to your account
   - Navigate to Account Settings > API Keys
   - Create a new API key with "Push new packages and package versions" scope
   - Copy the generated API key (keep it secure!)

3. **.NET SDK**: Ensure you have the .NET 10 SDK installed.

## Package Information

- **Package ID**: OpenDocx.NET
- **Version**: 1.1.8
- **Target Framework**: .NET 10.0
- **License**: MPL-2.0

## Publishing Steps

### Method 1: Using the Provided Script (Recommended)

1. Build the package:
   ```bash
   cd OpenDocx
   dotnet pack -c Release
   ```

2. Run the publish script:
   ```bash
   ./publish-nuget.sh YOUR_API_KEY_HERE
   ```

### Method 2: Manual Publishing

1. Build the package:
   ```bash
   cd OpenDocx
   dotnet pack -c Release
   ```

2. Push to NuGet:
   ```bash
   dotnet nuget push bin/Release/OpenDocx.NET.1.1.8.nupkg --api-key YOUR_API_KEY_HERE --source https://api.nuget.org/v3/index.json
   ```

## Package Contents

The generated package includes:

- **Main Assembly**: `OpenDocx.dll` - The core library
- **Symbol Package**: `OpenDocx.NET.1.1.8.snupkg` - Debug symbols for easier debugging
- **README**: The repository-root `README.md`, included as the canonical package guide
- **Dependencies**: All referenced NuGet packages are automatically included as dependencies

## Dependencies

The package has the following dependencies:

- DocumentFormat.OpenXml (>= 3.5.1)
- Newtonsoft.Json (>= 13.0.3)
- System.IO.Packaging (>= 10.0.2)
- SkiaSharp (>= 4.153.1)
- SkiaSharp.NativeAssets.Linux.NoDependencies (>= 4.153.1)

Note: Packages are created explicitly with `dotnet pack`; ordinary `dotnet build`
does not regenerate the OpenDocx package. OpenXmlPowerTools is bundled directly
into this package, so run the pack command above after changing either library.
Verify that the packaged `lib/net10.0/OpenXmlPowerTools.dll` matches the current
build before publishing.

## Version Management

To publish a new version:

1. Update the `<Version>` element in `OpenDocx.csproj`
2. Update the `<PackageReleaseNotes>` with changes
3. Rebuild and republish

## Post-Publication

After successful publication:

1. The package will be available at: https://www.nuget.org/packages/OpenDocx.NET/
2. It may take a few minutes to appear in search results
3. Users can install it with: `dotnet add package OpenDocx.NET`

## Troubleshooting

### Common Issues

1. **Package already exists**: If you're trying to push the same version again, increment the version number.

2. **API Key issues**: Ensure your API key has the correct permissions and hasn't expired.

3. **Build errors**: Make sure all dependencies are restored with `dotnet restore`.

4. **Package validation errors**: NuGet validates packages before accepting them. Check the error message for specific issues.

### Getting Help

If you encounter issues:

1. Check the [NuGet documentation](https://docs.microsoft.com/en-us/nuget/)
2. Visit the [OpenDocx GitHub repository](https://github.com/opendocx/opendocx-net)
3. Contact the project maintainers
