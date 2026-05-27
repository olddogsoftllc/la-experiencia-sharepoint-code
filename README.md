:

# The SharePoint Experience - Code Examples

Official code repository for the book **"The SharePoint Experience: Development with Microsoft Graph"** (La Experiencia SharePoint: Desarrollo con Microsoft Graph).

## 📚 About

This repository contains practical code examples in 5 programming languages:

- **C#** (.NET 6+)
- **PowerShell** (5.1 and 7.x)
- **JavaScript/Node.js** (18+)
- **Python** (3.9+)
- **Java** (17+)

## 📁 Structure

```
├── chapter-02-auth/              # Authentication and security
├── chapter-03-sites/             # Site management
├── chapter-04-documents/         # File handling
├── chapter-05-permissions/       # Permissions and sharing
├── chapter-06-metadata/            # Metadata and content types
├── chapter-07-automation/          # Webhooks and automation
├── complete-workflow/            # End-to-end workflow project
├── package.json                  # Node.js dependencies
├── requirements.txt              # Python dependencies
└── pom.xml                       # Java dependencies (Maven)
```

## 🚀 Quick Start

### Required Environment Variables

```bash
SP_TENANT_ID="your-tenant-id"
SP_CLIENT_ID="your-client-id"
SP_CLIENT_SECRET="your-client-secret"
SP_SITE_ID="your-site-id"
SP_DRIVE_ID="your-drive-id"
```

### Installation by Language

**JavaScript/Node.js:**
```bash
npm install
```

**Python:**
```bash
pip install -r requirements.txt
```

**Java:**
```bash
mvn clean install
```

**PowerShell:**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**C#:**
```bash
dotnet restore
```

## 📖 Usage

Each chapter contains standalone examples that you can run directly. Check the comments in each file for specific instructions.

## 📄 License

MIT License

---

**Note:** This repository contains code only. Complete documentation, guides, and detailed explanations are in the book.
