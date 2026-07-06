# Capítulo 6: Managed Metadata y Taxonomía - Ejemplos de Código

Este directorio contiene ejemplos prácticos para trabajar con el Servicio de Metadatos Administrados (Managed Metadata Service) de SharePoint Online mediante Microsoft Graph.

## 📁 Estructura de Carpetas

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

## 🎯 Temas Cubiertos

### 01 - Term Store

- Explorar el Term Store del tenant
- Listar grupos, conjuntos de términos y términos
- Crear estructuras de taxonomía
- Gestionar términos jerárquicos
- Sincronizar términos desde fuentes externas

### 02 - Content Types y Taxonomía

- Crear Content Types con columnas de metadatos
- Agregar columnas de taxonomía a bibliotecas
- Establecer valores de metadatos en documentos
- Buscar y filtrar por valores de taxonomía
- Mantenimiento de términos

## 📋 Requisitos Previos

### Para C#

- .NET 6.0 o superior
- Paquetes NuGet:
  - Microsoft.Graph >= 5.x
  - Azure.Identity >= 1.10.x

### Para PowerShell

- PowerShell 7.x
- Módulo PnP.PowerShell >= 2.x
- O Microsoft Graph PowerShell SDK

### Para JavaScript/Node.js

- Node.js >= 18.x
- Paquetes npm:
  - @microsoft/microsoft-graph-client
  - @azure/identity

### Para Python

- Python >= 3.9
- Paquetes:
  - msgraph-sdk
  - azure-identity

### Para Java

- Java 17 o superior
- Maven o Gradle
- Dependencias:
  - com.microsoft.graph:microsoft-graph
  - com.azure:azure-identity

## 🔐 Configuración

### Permisos Requeridos de Microsoft Graph

Para ejecutar estos ejemplos se necesitan los siguientes permisos:

| Permiso | Tipo | Descripción |
|---------|------|-------------|
| `TermStore.Read.All` | Delegado | Leer términos y conjuntos |
| `TermStore.ReadWrite.All` | Delegado | Crear/modificar términos |
| `Sites.Read.All` | Delegado | Leer sitios |
| `Sites.FullControl.All` | Delegado | Gestionar columnas y Content Types |

## 🚀 Ejecución

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

## 📚 Referencias

- [Documentación de Microsoft Graph - Term Store](https://docs.microsoft.com/en-us/graph/api/resources/termstore-store)
- [SharePoint Managed Metadata](https://docs.microsoft.com/en-us/sharepoint/managed-metadata)
- [Content Types en SharePoint](https://docs.microsoft.com/en-us/sharepoint/governance/content-types-overview)

## 💡 Notas Importantes

1. **Requiere cuenta con permisos de administrador** para gestionar el Term Store
2. **El Term Store es de nivel tenant**, los cambios afectan a toda la organización
3. **Las columnas de metadatos** deben referenciar conjuntos de términos existentes
4. **Los términos tienen IDs únicos** que se usan para establecer valores, no los labels

## 📝 Ejercicios Sugeridos

1. Mapear todo el Term Store de tu tenant y exportar a CSV
2. Crear una taxonomía de proyectos con departamentos y equipos
3. Agregar columnas de metadatos a una biblioteca y probar la clasificación
4. Crear un Content Type con múltiples campos de taxonomía
5. Implementar búsqueda filtrada por valores de taxonomía
6. Crear un script de sincronización desde Excel/CSV
