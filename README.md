## PowerPoint Remote Control

### Description

A Windows application that lets you control a PowerPoint presentation remotely from a mobile device via **Bluetooth**.
No internet connection, no cloud infrastructure, no accounts — just pair and present.

## Technologies

- [.NET Framework 4.8.1](https://dotnet.microsoft.com/)
- [C#](https://docs.microsoft.com/en-us/dotnet/csharp/)
- [InTheHand.Net.Bluetooth (32feet.NET)](https://github.com/inthehand/32feet) — Bluetooth RFCOMM
- [Microsoft.Office.Interop.PowerPoint](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)

## Architecture

```
[Mobile App] --Bluetooth RFCOMM--> [PC: This App] --> [PowerPoint]
```

The PC acts as a Bluetooth RFCOMM server using the **Serial Port Profile (SPP)** UUID.
The mobile client pairs with the PC and sends plain-text commands over the Bluetooth socket.

## Why Bluetooth?

| | SQS (original) | Redis Pub/Sub | Bluetooth (current) |
|---|---|---|---|
| Latency | Up to 20s | < 1ms | 2–5ms |
| Internet required | Yes | Yes | **No** |
| Infrastructure | AWS account | Redis server | **None** |
| Range | Unlimited | Unlimited | ~10m |
| Setup | Complex | Moderate | **Pair & go** |

## Supported Commands

Send these as plain-text lines (terminated with `\n`) over the Bluetooth socket:

| Command | Action |
|---|---|
| `next` | Advance to next slide |
| `previous` | Go back to previous slide |
| `first` | Jump to first slide |
| `last` | Jump to last slide |

## Bluetooth Service Info

| Property | Value |
|---|---|
| Profile | RFCOMM / Serial Port Profile (SPP) |
| UUID | `00001101-0000-1000-8000-00805f9b34fb` |
| Service Name | `PowerPoint Remote Control` |

## How to Use

1. Clone this repository
2. Open the solution in Visual Studio
3. Restore NuGet packages (`InTheHand.Net.Bluetooth`)
4. Build and run the application
5. **Pair** your mobile device with the PC via Windows Bluetooth settings
6. On the mobile app, connect to the service named **"PowerPoint Remote Control"**
7. Send commands as plain-text lines over the socket

### Testing from another device

Any app that can connect to a Bluetooth SPP service works as a client — e.g.:
- **Android**: "Serial Bluetooth Terminal" (Play Store)
- **iOS**: "Bluetooth Terminal" apps that support SPP/RFCOMM

## How to Contribute

1. Fork the project
2. Create a feature branch
3. Submit a pull request
