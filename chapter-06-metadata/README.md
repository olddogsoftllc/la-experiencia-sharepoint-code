# Chapter 6: Managed Metadata and Taxonomy - Code Samples

This directory contains practical examples for working with the SharePoint Online Managed Metadata Service via Microsoft Graph.

## 📁 Folder Structure

```
capitulo-06-managed-metadata/
├── csharp/
│   ├── 01-termstore/
│   │   ├── TermStoreExplorer.cs
│   │   ├── TaxonomyManager.cs
│   │   └── MetadataColumnManager.cs
│   └── 02-contenttypes-taxonomy/
├── powershell/
│   ├── 01-termstore/
│   │   ├── Explore-TermStore.ps1
│   │   └── New-TaxonomyStructure.ps1
│   └── 02-contenttypes-taxonomy/
├── javascript/
│   ├── 01-termstore/
│   │   ├── exploreTermStore.js
│   │   └── taxonomyManager.js
│   └── 02-contenttypes-taxonomy/
├── python/
│   ├── 01-termstore/
│   │   ├── explore_term_store.py
│   │   └── taxonomy_manager.py
│   └── 02-contenttypes-taxonomy/
└── java/
    ├── 01-termstore/
    │   ├── TermStoreExplorer.java
    │   └── TaxonomyManager.java
    └── 02-contenttypes-taxonomy/
```

## 🎯 Topics Covered

### 01 - Term Store

- Explore the tenant's Term Store
- List groups, term sets, and terms
- Create taxonomy structures
- Manage hierarchical terms
- Sync terms from external sources

### 02 - Content Types and Taxonomy

- Create Content Types with metadata columns
- Add taxonomy columns to libraries
- Set metadata values on documents
- Search and filter by taxonomy values
- Term maintenance

## 📋 Prerequisites

### For C#

- .NET 6.0 or higher
- NuGet packages:
  - Microsoft.Graph >= 5.x
  - Azure.Identity >= 1.10.x

### For PowerShell

- PowerShell 7.x
- PnP.PowerShell module >= 2.x
- Or Microsoft Graph PowerShell SDK

### For JavaScript/Node.js

- Node.js >= 18.x
- npm packages:
  - @microsoft/microsoft-graph-client
  - @azure/identity

### For Python

- Python >= 3.9
- Packages:
  - msgraph-sdk
  - azure-identity

### For Java

- Java 17 or higher
- Maven or Gradle
- Dependencies:
  - com.microsoft.graph:microsoft-graph
  - com.azure:azure-identity

## 🔐 Configuration

### Required Microsoft Graph Permissions

The following permissions are needed to run these examples:

| Permission | Type | Description |
|---------|------|-------------|
| `TermStore.Read.All` | Delegated | Read terms and term sets |
| `TermStore.ReadWrite.All` | Delegated | Create/modify terms |
| `Sites.Read.All` | Delegated | Read sites |
| `Sites.FullControl.All` | Delegated | Manage columns and Content Types |

## 🚀 Execution

### C# (Interactive)

```bash
cd csharp/01-termstore
dotnet run -- TermStoreExplorer
```

### PowerShell

```powershell
# Ejecutar con autenticación interactiva
.\Explore-TermStore.ps1 -SiteUrl "https://tu-tenant.sharepoint.com"
```

### JavaScript (Node.js)

```bash
cd javascript/01-termstore
npm install
node exploreTermStore.js
```

### Python

```bash
cd python/01-termstore
pip install -r requirements.txt
python explore_term_store.py
```

### Java

```bash
cd java/01-termstore
mvn compile exec:java -Dexec.mainClass="TermStoreExplorer"
```

## 📚 References

- [Microsoft Graph documentation - Term Store](https://docs.microsoft.com/en-us/graph/api/resources/termstore-store)
- [SharePoint Managed Metadata](https://docs.microsoft.com/en-us/sharepoint/managed-metadata)
- [Content Types in SharePoint](https://docs.microsoft.com/en-us/sharepoint/governance/content-types-overview)

## 💡 Important Notes

1. **Requires an account with administrator permissions** to manage the Term Store
2. **The Term Store is tenant-level**, changes affect the entire organization
3. **Metadata columns** must reference existing term sets
4. **Terms have unique IDs** that are used to set values, not the labels

## 📝 Suggested Exercises

1. Map your entire tenant Term Store and export it to CSV
2. Create a project taxonomy with departments and teams
3. Add metadata columns to a library and try classification
4. Create a Content Type with multiple taxonomy fields
5. Implement filtered search by taxonomy values
6. Create a sync script from Excel/CSV