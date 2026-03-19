## PowerPoint Remote Control

### Description

A Windows application that lets you control a PowerPoint presentation remotely from a mobile device.
Uses **Redis Pub/Sub** for real-time, low-latency message delivery (sub-millisecond vs ~20s with SQS).

## Technologies

- [.NET Framework 4.8.1](https://dotnet.microsoft.com/)
- [C#](https://docs.microsoft.com/en-us/dotnet/csharp/)
- [Redis](https://redis.io/) via [StackExchange.Redis](https://github.com/StackExchange/StackExchange.Redis)
- [Microsoft.Office.Interop.PowerPoint](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)

## Why Redis instead of SQS?

| | SQS (old) | Redis Pub/Sub (new) |
|---|---|---|
| Latency | Up to 20 seconds (long polling) | < 1ms (push-based) |
| Model | Pull (polling) | Push (event-driven) |
| Infrastructure | AWS account required | Local or Redis Cloud free tier |
| Cost | Per request pricing | Free self-hosted |

## Supported Commands

| Command | Action |
|---|---|
| `next` | Advance to next slide |
| `previous` | Go back to previous slide |
| `first` | Jump to first slide |
| `last` | Jump to last slide |

## Environment Variables

| Variable | Description | Default |
|---|---|---|
| `REDIS_CONNECTION_STRING` | Redis connection string | `localhost:6379` |
| `REDIS_CHANNEL` | Pub/Sub channel name | `powerpoint-control` |

## How to Use

1. Clone this repository
2. Install Redis locally or create a free instance at [Redis Cloud](https://redis.com/try-free/)
3. Open the solution in Visual Studio
4. Restore NuGet packages
5. Build and run the application
6. Clone the [PowerPoint Remote Control Client](https://github.com/thethiago27/power_point_control_client) and configure it to publish to the same Redis channel

### Quick Start with Local Redis

```bash
# Start Redis (Docker)
docker run -d -p 6379:6379 redis:alpine

# Set environment variables (PowerShell)
$env:REDIS_CONNECTION_STRING = "localhost:6379"
$env:REDIS_CHANNEL = "powerpoint-control"

# Run the app, then test from another terminal:
docker exec -it <container> redis-cli PUBLISH powerpoint-control next
```

## How to Contribute

1. Fork the project
2. Create a feature branch
3. Submit a pull request
