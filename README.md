# Simple Send

Source code for Simple Send, a Windows Forms front end for the
`net send` messenger command.

## What it does

1. Select recipients from a list of computer or user names, adding or
   removing names manually and saving the list.
2. Type the message text.
3. Click **Send**. The message is split into chunks and delivered to
   each recipient with `net send`.

`net send` relies on the Windows Messenger service, which was removed
in Windows Vista, so the tool only works on older Windows versions.

## Tech

- Visual Basic .NET (Windows Forms)

## Building

Open the `.sln` file in Visual Studio and build it. The project was
upgraded to the Visual Studio 2005 format (.NET Framework 2.0) and targets
Windows only.

The solution file is `SimpleSend.sln`.

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE).
