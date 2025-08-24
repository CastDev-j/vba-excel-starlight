# Guía de Contribución

¡Gracias por tu interés en contribuir a la documentación de VBA para Excel! Esta guía te ayudará a participar de manera efectiva.

## Formas de Contribuir

### Reportar Problemas

- Errores en el código de ejemplos
- Typos o errores gramaticales
- Enlaces rotos
- Sugerencias de contenido nuevo

### Mejorar Contenido Existente

- Clarificar explicaciones confusas
- Añadir ejemplos adicionales
- Mejorar código existente
- Actualizar información desactualizada

### Crear Contenido Nuevo

- Nuevos ejemplos prácticos
- Casos de uso específicos
- Tutoriales paso a paso

## Proceso de Contribución

### 1. Fork y Clone del Repositorio

```bash
git clone https://github.com/tu-usuario/vba-excel-docs.git
cd vba-excel-docs
git remote add upstream https://github.com/usuario-original/vba-excel-docs.git
```

### 2. Crear una Rama

```bash
git checkout -b feature/nuevo-ejemplo
# o
git checkout -b fix/corregir-error
```

### 3. Realizar Cambios

- Sigue las convenciones de estilo
- Prueba tu código antes de enviarlo
- Asegúrate de que los enlaces funcionen

### 4. Enviar Cambios

```bash
git add .
git commit -m "Descripción clara del cambio"
git push origin feature/nuevo-ejemplo
```

### 5. Crear Pull Request

1. Ve a tu fork en GitHub
2. Crea un Pull Request hacia la rama main
3. Describe claramente los cambios realizados

## Convenciones de Estilo

### Estructura de Archivos

- Usa `kebab-case` para nombres de archivos
- Organiza el contenido por categorías lógicas
- Incluye frontmatter completo en cada archivo

### Formato de Código VBA

**Recomendado:**

```vba
' Usar nombres descriptivos y comentarios claros
Sub GenerarReporte()
    Dim ultimaFila As Long
    Dim rangoVentas As Range

    ' Obtener la última fila con datos
    ultimaFila = Cells(Rows.Count, 1).End(xlUp).Row
    Set rangoVentas = Range("A1:D" & ultimaFila)

    ' Manejar errores
    On Error GoTo ErrorHandler

    rangoVentas.Sort Key1:=Range("A1")
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
```

**Evitar:**

```vba
' Nombres poco claros, sin comentarios
sub proceso()
    dim x, y
    x = cells(1000000, 1).end(-4162).row
    set y = range("A1:D" & x)
    y.sort range("A1")
end sub
```

### Estructura de Documentos

````markdown
---
title: "Título Descriptivo"
description: "Descripción concisa"
---

# Título Principal

Introducción breve explicando el contenido.

## Conceptos Importantes

Explicación de conceptos clave.

## Ejemplo Práctico

### Código

\```vba
' Código bien comentado
Sub Ejemplo()
' Explicar cada paso importante
Dim variable As String
variable = "ejemplo"
End Sub
\```

### Explicación

Descripción paso a paso del código.

## Recursos Relacionados

- [Enlace a documentación relacionada](/ruta/relacionada)
````

## Tipos de Contribución

### Reportar Errores

Al reportar un error, incluye:

- Descripción clara del problema
- Pasos para reproducir el error
- Comportamiento esperado vs actual
- Versión de Excel y sistema operativo si es relevante

### Sugerir Mejoras

Para sugerencias, proporciona:

- Descripción del problema actual
- Solución propuesta
- Beneficios de la mejora
- Alternativas consideradas

### Nuevo Contenido

Al proponer contenido nuevo:

- Define la audiencia objetivo (principiante/intermedio/avanzado)
- Explica qué problema resuelve
- Proporciona un outline del contenido

## Convenciones de Commits

Usa mensajes descriptivos y claros:

- `feat: agregar ejemplo de dashboard interactivo`
- `fix: corregir error en bucle infinito`
- `docs: mejorar explicación del objeto Range`
- `style: corregir formato del código`

## Template para Pull Requests

```markdown
## Descripción

Breve descripción de los cambios realizados.

## Tipo de cambio

- [ ] Corrección de error
- [ ] Nueva funcionalidad
- [ ] Mejora de documentación
- [ ] Mejora de estilo/formato

## Testing

- [ ] Código VBA probado y funcional
- [ ] Enlaces verificados
- [ ] Revisión ortográfica completada

## Checklist

- [ ] Los cambios siguen las convenciones de estilo
- [ ] He revisado mis cambios
- [ ] El código está comentado apropiadamente
- [ ] No se introducen nuevos errores
```

## Obtener Ayuda

Si necesitas ayuda:

1. Revisa los issues existentes en GitHub
2. Consulta la documentación existente
3. Crea un nuevo issue con tu pregunta
4. Proporciona contexto suficiente (código, errores, pasos)

## Código de Conducta

Mantén un ambiente respetuoso y profesional:

- Usa lenguaje constructivo y respetuoso
- Acepta críticas de manera profesional
- Ayuda a otros contribuidores cuando sea posible
- Enfócate en mejorar la documentación para toda la comunidad

## Reconocimiento

Todos los contribuidores serán reconocidos en la página de créditos del proyecto y en las notas de lanzamiento cuando corresponda.

Gracias por contribuir a hacer esta documentación mejor para toda la comunidad.
