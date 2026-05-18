# OneNote Study Planner

A .NET console application that automatically creates structured OneNote study pages using Microsoft Graph API.

## Features

- Authenticate with Microsoft account
- Read OneNote notebooks and sections
- Import roadmap from JSON
- Automatically create OneNote pages
- Skip already existing pages
- Pagination support for Graph API

## Tech Stack

- .NET 8
- Microsoft Graph API
- Azure Identity
- OneNote API

## Setup

1. Create Azure App Registration
2. Add Graph Permissions:
   - User.Read
   - Notes.ReadWrite
3. Update appsettings.json
4. Run application

## Example roadmap.json

```json
[
  {
    "day": 1,
    "title": "C# Fundamentals",
    "topics": [
      "Variables",
      "Loops"
    ]
  }
]
```

## Run

```bash
dotnet run
```