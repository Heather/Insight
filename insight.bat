@echo off
cls
SET EnableNuGetPackageRestore=true
if %PROCESSOR_ARCHITECTURE%==x86 (
         set MSBUILD="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
) else ( set MSBUILD="%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
)
if not exist .nuget\nuget.exe %MSBUILD% .nuget\nuget.targets /t:CheckPrerequisites
if not exist packages\FAKE\tools\Fake.exe ( 
	echo Downloading FAKE...
	".nuget\NuGet.exe" "install" "FAKE" "-OutputDirectory" "packages" "-ExcludeVersion" "-Prerelease"
)
if not exist packages\ModernUI\lib\net40\MetroFramework.dll ( 
	echo Downloading ModernUI...
	".nuget\NuGet.exe" "install" "ModernUI" "-OutputDirectory" "packages" "-ExcludeVersion" "-Prerelease"
)
if not exist packages\FSharp.Charting\lib\net40\FSharp.Charting.dll ( 
	echo Downloading FSharp.Charting...
	".nuget\NuGet.exe" "install" "FSharp.Charting" "-OutputDirectory" "packages" "-ExcludeVersion" "-Prerelease"
)
if not exist packages\NPOI\lib\net40\NPOI.dll ( 
	echo Downloading NPOI...
	".nuget\NuGet.exe" "install" "NPOI" "-OutputDirectory" "packages" "-ExcludeVersion" "-Prerelease"
)

start packages\FAKE\tools\Fsi.exe insight.fsx