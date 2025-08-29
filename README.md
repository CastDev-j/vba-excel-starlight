# VBA para Excel - Documentación

Guía completa y práctica de VBA para Excel con ejemplos reales.

## Instalación

```bash
npm install
npm run dev
```

## Estructura

```
src/content/docs/
├── getting-started/     # Primeros pasos
├── syntax/             # Sintaxis básica
├── objects/            # Objetos de Excel
├── examples/           # Ejemplos por nivel
├── reference/          # Referencia API
└── troubleshooting/    # Solución de problemas
```

## Contribuir

1. Fork el proyecto
2. Crea tu rama: `git checkout -b mi-contribucion`
3. Commit: `git commit -m 'Agregar nuevo ejemplo'`
4. Push: `git push origin mi-contribucion`
5. Abre un Pull Request

## Comandos

```bash
npm run dev         # Servidor de desarrollo
npm run build       # Build para producción
npm run preview     # Preview del build
```

## Licencia

Esta documentación se basa en la documentación oficial de VBA (Visual Basic for Applications) de Microsoft y se distribuye bajo la licencia Creative Commons Attribution 4.0 International (CC BY‑4.0).

## Progreso

src/content/docs/
├── index.mdx                           
├── getting-started/
│   ├── introduction.mdx                ✓
│   ├── setup.mdx                       ✓
│   ├── hello-world.mdx                 ✓
│   └── concepts.mdx                    ✓
├── syntax/
│   ├── variables-datatypes.mdx         ✓
│   ├── operators.mdx                   ✓
│   ├── control-structures.mdx          ✓
│   ├── procedures-functions.mdx        ✓
│   ├── arrays.mdx                      ✓
│   └── error-handling.mdx              ✓
├── objects/
│   ├── application.mdx                 ✓
│   ├── workbook.mdx                    ✓
│   ├── worksheet.mdx                   ✓
│   ├── range.mdx                       ✓
│   ├── cells-ranges.mdx                ✓
│   ├── charts.mdx                      ✓
│   ├── pivottables.mdx                 ✓
│   ├── listobjects.mdx                 ✓
│   ├── names.mdx                       ✓
│   ├── userforms.mdx                   ✓
│   ├── controls.mdx                    ✓
│   ├── commandbars.mdx                 ✓
│   └── events.mdx                      ✓
├── techniques/
│   ├── cell-operations.mdx             # Leer/escribir celdas eficientemente
│   ├── loops-iterations.mdx            # For Each, optimización de bucles
│   ├── filtering-sorting.mdx           # AutoFilter, Advanced Filter, Sort
│   ├── find-replace.mdx                # Find, FindNext, Replace methods
│   ├── conditional-formatting.mdx      # Formateo condicional programático
│   ├── data-validation.mdx             # Validation objects y reglas
│   └── import-export.mdx               # CSV, TXT, otros formatos
├── examples/
│   ├── beginner/
│   │   ├── auto-format.mdx             # Formatear automáticamente
│   │   ├── calculator.mdx              # Calculadora simple
│   │   ├── report-generator.mdx        # Generar reportes básicos
│   │   └── data-cleanup.mdx            # Limpiar y ordenar datos
│   ├── intermediate/
│   │   ├── dashboard.mdx               # Dashboard interactivo
│   │   ├── inventory-system.mdx        # Sistema de inventario
│   │   ├── sales-analysis.mdx          # Análisis de ventas
│   │   ├── dynamic-forms.mdx           # Formularios que cambian
│   │   └── database-connection.mdx     # Conectar a SQL Server/Access
│   └── advanced/
│       ├── mini-erp.mdx                # Sistema ERP básico
│       ├── web-scraping.mdx            # Extraer datos de web
│       ├── office-integration.mdx      # Integrar Word/PowerPoint
│       ├── custom-addins.mdx           # Crear add-ins personalizados
│       └── web-apis.mdx                # Consumir APIs REST
├── best-practices/
│   ├── clean-code.mdx                  # Código limpio y legible
│   ├── performance.mdx                 # Optimización de rendimiento
│   ├── debugging-testing.mdx           # Debug y testing strategies
│   ├── security.mdx                    # Seguridad y validación
│   ├── documentation.mdx               # Documentar código VBA
│   └── memory-management.mdx           # Manejo eficiente de memoria
├── use-cases/
│   ├── accounting-automation.mdx       # Automatización contable
│   ├── project-management.mdx          # Gestión de proyectos
│   ├── financial-analysis.mdx          # Análisis financiero
│   ├── quality-control.mdx             # Control de calidad
│   ├── human-resources.mdx             # Recursos humanos
│   └── logistics.mdx                   # Logística y supply chain
├── tools/
│   ├── ides-editors.mdx                # Editores y IDEs para VBA
│   ├── debugging.mdx                   # Herramientas de debugging
│   ├── testing.mdx                     # Testing frameworks
│   ├── version-control.mdx             # Git para proyectos VBA
│   ├── libraries.mdx                   # Librerías útiles
│   └── online-resources.mdx            # Recursos online
├── reference/
│   ├── application-methods.mdx         # Application methods A-Z
│   ├── application-properties.mdx      # Application properties A-Z
│   ├── workbook-methods.mdx            # Workbook methods A-Z
│   ├── workbook-properties.mdx         # Workbook properties A-Z
│   ├── worksheet-methods.mdx           # Worksheet methods A-Z
│   ├── worksheet-properties.mdx        # Worksheet properties A-Z
│   ├── range-methods.mdx               # Range methods A-Z
│   ├── range-properties.mdx            # Range properties A-Z
│   ├── chart-methods.mdx               # Chart methods A-Z
│   ├── chart-properties.mdx            # Chart properties A-Z
│   ├── pivottable-methods.mdx          # PivotTable methods A-Z
│   ├── pivottable-properties.mdx       # PivotTable properties A-Z
│   ├── vba-functions.mdx               # Built-in VBA functions
│   ├── excel-functions.mdx             # WorksheetFunction methods
│   ├── constants-enums.mdx             # XlConstants y enumeraciones
│   ├── data-types.mdx                  # Todos los data types
│   └── error-codes.mdx                 # Códigos de error comunes
├── troubleshooting/
│   ├── common-errors.mdx               # Errores más frecuentes
│   ├── faq.mdx                         # Preguntas frecuentes
│   ├── performance-issues.mdx          # Problemas de rendimiento
│   ├── version-compatibility.mdx       # Compatibilidad entre versiones
│   └── code-migration.mdx              # Migrar código legacy
└── appendices/
    ├── glossary.mdx                    # Glosario de términos VBA
    ├── keyboard-shortcuts.mdx          # Atajos de teclado VBE
    ├── error-codes.mdx                 # Lista completa códigos error
    ├── excel-functions.mdx             # Funciones Excel en VBA
    ├── constants-enums.mdx             # Constantes y enumeraciones
    └── additional-resources.mdx        # Libros, cursos, comunidades