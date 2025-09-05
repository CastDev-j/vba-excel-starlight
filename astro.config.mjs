// @ts-check
// astro.config.mjs
import { defineConfig } from 'astro/config';
import starlight from '@astrojs/starlight';

import astroExpressiveCode from 'astro-expressive-code'
import tailwindcss from '@tailwindcss/vite';
import fs from "node:fs";

const myLang = JSON.parse(fs.readFileSync('./src/assets/vba.json', 'utf-8'));

export default defineConfig({
	integrations: [
		astroExpressiveCode({
			shiki: {
				langs: [
					myLang
				],
			},

		}),
		starlight({
			title: 'VBA Excel',
			description: 'Guía completa de VBA para Excel con ejemplos prácticos y referencia detallada',

			// Logo personalizado (opcional)
			logo: {
				dark: './src/assets/dark-excel-logo.svg',
				light: './src/assets/light-excel-logo.svg',
				replacesTitle: false,
			},

			defaultLocale: 'es',

			customCss: [
				'./src/styles/global.css',
			],

			// Configuración del sidebar - ESTRUCTURA COMPLETA
			sidebar: [
				// SECCIÓN 1: INTRODUCCIÓN Y FUNDAMENTOS
				{
					label: 'Primeros Pasos',
					collapsed: true,
					items: [
						{ label: 'Introducción a VBA', slug: 'getting-started/introduction' },
						{ label: 'Configuración del Entorno', slug: 'getting-started/setup' },
						{ label: 'Tu Primer Programa', slug: 'getting-started/hello-world' },
						{ label: 'Conceptos Fundamentales', slug: 'getting-started/concepts' },
					],
				},

				// SECCIÓN 2: SINTAXIS Y FUNDAMENTOS DEL LENGUAJE
				{
					label: 'Sintaxis Básica',
					collapsed: true,
					items: [
						{ label: 'Variables y Tipos de Datos', slug: 'syntax/variables-datatypes' },
						{ label: 'Operadores', slug: 'syntax/operators' },
						{ label: 'Estructuras de Control', slug: 'syntax/control-structures' },
						{ label: 'Procedimientos y Funciones', slug: 'syntax/procedures-functions' },
						{ label: 'Matrices (Arrays)', slug: 'syntax/arrays' },
						{ label: 'Manejo de Errores', slug: 'syntax/error-handling' },
					],
				},

				// SECCIÓN 3: MODELO DE OBJETOS DE EXCEL
				{
					label: 'Modelo de Objetos Excel',
					collapsed: true,
					items: [
						// Objetos principales
						{
							label: 'Objetos Fundamentales',
							collapsed: true,
							items: [
								{ label: 'Application Object', slug: 'objects/application' },
								{ label: 'Workbook Object', slug: 'objects/workbook' },
								{ label: 'Worksheet Object', slug: 'objects/worksheet' },
								{ label: 'Range Object', slug: 'objects/range' },
								{ label: 'Cells y Ranges', slug: 'objects/cells-ranges' },
							],
						},
						// Objetos de datos
						{
							label: 'Objetos de Datos',
							collapsed: true,
							items: [
								{ label: 'Chart Objects', slug: 'objects/charts' },
								{ label: 'PivotTable Objects', slug: 'objects/pivottables' },
								{ label: 'ListObjects (Tablas)', slug: 'objects/listobjects' },
								{ label: 'Name Objects', slug: 'objects/names' },
							],
						},
						// Objetos de interfaz
						{
							label: 'Objetos de Interfaz',
							collapsed: true,
							items: [
								{ label: 'UserForms', slug: 'objects/userforms' },
								{ label: 'Controls', slug: 'objects/controls' },
								{ label: 'CommandBars', slug: 'objects/commandbars' },
								{ label: 'Events', slug: 'objects/events' },
							],
						},
					],
				},

				// SECCIÓN 4: TÉCNICAS Y PATRONES COMUNES
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Técnicas Esenciales',
					collapsed: true,
					items: [
						{ label: 'Lectura y Escritura de Celdas', slug: 'techniques/cell-operations' },
						{ label: 'Bucles y Iteraciones', slug: 'techniques/loops-iterations' },
						{ label: 'Filtrado y Ordenación', slug: 'techniques/filtering-sorting' },
						{ label: 'Búsqueda y Reemplazo', slug: 'techniques/find-replace' },
						{ label: 'Formateo Condicional', slug: 'techniques/conditional-formatting' },
						{ label: 'Validación de Datos', slug: 'techniques/data-validation' },
						{ label: 'Importación/Exportación', slug: 'techniques/import-export' },
					],
				},

				// SECCIÓN 5: EJEMPLOS PRÁCTICOS POR CATEGORÍA
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Ejemplos Prácticos',
					collapsed: true,
					items: [
						// Ejemplos básicos
						{
							label: 'Nivel Principiante',
							collapsed: true,
							items: [
								{ label: 'Automatizar Formato', slug: 'examples/beginner/auto-format' },
								{ label: 'Calculadora Simple', slug: 'examples/beginner/calculator' },
								{ label: 'Generador de Reportes', slug: 'examples/beginner/report-generator' },
								{ label: 'Limpieza de Datos', slug: 'examples/beginner/data-cleanup' },
							],
						},
						// Ejemplos intermedios
						{
							label: 'Nivel Intermedio',
							collapsed: true,
							items: [
								{ label: 'Dashboard Interactivo', slug: 'examples/intermediate/dashboard' },
								{ label: 'Sistema de Inventario', slug: 'examples/intermediate/inventory-system' },
								{ label: 'Análisis de Ventas', slug: 'examples/intermediate/sales-analysis' },
								{ label: 'Formularios Dinámicos', slug: 'examples/intermediate/dynamic-forms' },
								{ label: 'Conexión a Base de Datos', slug: 'examples/intermediate/database-connection' },
							],
						},
						// Ejemplos avanzados
						{
							label: 'Nivel Avanzado',
							collapsed: true,
							items: [
								{ label: 'Sistema ERP Mini', slug: 'examples/advanced/mini-erp' },
								{ label: 'Web Scraping', slug: 'examples/advanced/web-scraping' },
								{ label: 'Integración Office', slug: 'examples/advanced/office-integration' },
								{ label: 'Add-ins Personalizados', slug: 'examples/advanced/custom-addins' },
								{ label: 'APIs y Servicios Web', slug: 'examples/advanced/web-apis' },
							],
						},
					],
				},

				// SECCIÓN 6: MEJORES PRÁCTICAS Y OPTIMIZACIÓN
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Mejores Prácticas',
					collapsed: true,
					items: [
						{ label: 'Código Limpio en VBA', slug: 'best-practices/clean-code' },
						{ label: 'Optimización de Rendimiento', slug: 'best-practices/performance' },
						{ label: 'Debugging y Testing', slug: 'best-practices/debugging-testing' },
						{ label: 'Seguridad y Validación', slug: 'best-practices/security' },
						{ label: 'Documentación de Código', slug: 'best-practices/documentation' },
						{ label: 'Manejo de Memoria', slug: 'best-practices/memory-management' },
					],
				},

				// SECCIÓN 7: CASOS DE USO ESPECÍFICOS
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Casos de Uso',
					collapsed: true,
					items: [
						{ label: 'Automatización Contable', slug: 'use-cases/accounting-automation' },
						{ label: 'Gestión de Proyectos', slug: 'use-cases/project-management' },
						{ label: 'Análisis Financiero', slug: 'use-cases/financial-analysis' },
						{ label: 'Control de Calidad', slug: 'use-cases/quality-control' },
						{ label: 'Recursos Humanos', slug: 'use-cases/human-resources' },
						{ label: 'Logística y Supply Chain', slug: 'use-cases/logistics' },
					],
				},

				// SECCIÓN 8: HERRAMIENTAS Y RECURSOS
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Herramientas y Recursos',
					collapsed: true,
					items: [
						{ label: 'IDEs y Editores', slug: 'tools/ides-editors' },
						{ label: 'Debugging Tools', slug: 'tools/debugging' },
						{ label: 'Testing Frameworks', slug: 'tools/testing' },
						{ label: 'Version Control', slug: 'tools/version-control' },
						{ label: 'Librerías Útiles', slug: 'tools/libraries' },
						{ label: 'Recursos Online', slug: 'tools/online-resources' },
					],
				},

				// SECCIÓN 9: REFERENCIA COMPLETA
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Referencia API',
					collapsed: true,
					// Autogenerar desde directorio reference/
					autogenerate: {
						directory: 'reference',
						collapsed: true,
					},
				},

				// SECCIÓN 10: TROUBLESHOOTING Y FAQ
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Ayuda y Soluciones',
					collapsed: true,
					items: [
						{ label: 'Errores Comunes', slug: 'troubleshooting/common-errors' },
						{ label: 'FAQ - Preguntas Frecuentes', slug: 'troubleshooting/faq' },
						{ label: 'Problemas de Rendimiento', slug: 'troubleshooting/performance-issues' },
						{ label: 'Compatibilidad Versiones', slug: 'troubleshooting/version-compatibility' },
						{ label: 'Migración de Código', slug: 'troubleshooting/code-migration' },
					],
				},

				// SECCIÓN 11: APÉNDICES
				{
					badge: { text: 'En Fila', variant: 'caution' },
					label: 'Apéndices',
					collapsed: true,
					items: [
						{ label: 'Glosario de Términos', slug: 'appendices/glossary' },
						{ label: 'Atajos de Teclado', slug: 'appendices/keyboard-shortcuts' },
						{ label: 'Códigos de Error', slug: 'appendices/error-codes' },
						{ label: 'Funciones de Excel en VBA', slug: 'appendices/excel-functions' },
						{ label: 'Constantes y Enumeraciones', slug: 'appendices/constants-enums' },
						{ label: 'Recursos Adicionales', slug: 'appendices/additional-resources' },
					],
				},
			],

			// Configuración adicional
			editLink: {
				baseUrl: 'https://github.com/CastDev-j/vba-excel-starlight/tree/maincat',
			},

			// Configuración de búsqueda (requiere configuración adicional)
			pagefind: true,

			// Componentes personalizados para la página de inicio
			components: {
				// Puedes sobrescribir componentes por defecto
			},

			// Configuración para el archivo head personalizado
			head: [
				{
					tag: 'meta',
					attrs: {
						property: 'og:title',
						content: 'VBA Excel',
					},
				},
				{
					tag: 'meta',
					attrs: {
						property: 'og:description',
						content: 'La guía más completa de VBA para Excel con ejemplos prácticos',
					},
				},
				{
					tag: 'script',
					attrs: { type: 'module' },
					content: `
        import { inject } from '@vercel/analytics';
        inject();
      `
				}
			],

			// Configuración de navegación adicional
			lastUpdated: true,
			pagination: true,
		}),
	],

	vite: {
		plugins: [tailwindcss()],
	},
});