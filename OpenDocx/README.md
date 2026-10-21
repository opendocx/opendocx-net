# OpenDocx.NET

OpenDocx.NET translates Word templates containing text-delimited `{[fields]}`
into OpenXmlPowerTools templates. The host application supplies its own expression
evaluator and data; OpenDocx handles template provisioning and exposes OXPT
assembly through .NET APIs.

The [repository README](../README.md) is the canonical guide, including a working
C# example, the provisioning/assembly boundary, and build instructions. That same
README is included in the NuGet package.

See [PUBLISHING.md](PUBLISHING.md) for package-building and publishing instructions.
