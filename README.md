# Unificación Data Pastas

## Descripción General

**Unificación Data Pastas** es una aplicación web desarrollada con **Google Apps Script** cuyo propósito es centralizar, estandarizar y optimizar la gestión de reportes operativos dentro de la organización. El sistema permite registrar, consultar y dar seguimiento a distintos tipos de reportes generados en planta, integrando en una sola solución formularios web, almacenamiento estructurado de datos, evidencias en imágenes y notificaciones automáticas por correo electrónico.

El proyecto surge como una necesidad de **unificar múltiples flujos de información dispersos**, mejorar la trazabilidad de los reportes y facilitar la toma de decisiones mediante información consolidada y accesible.

---

## Alcance del Sistema

El sistema cubre el ciclo completo de gestión de reportes operativos:

- Registro de reportes desde una interfaz web.
- Clasificación por tipo de reporte.
- Almacenamiento estructurado de información.
- Asociación de evidencias gráficas.
- Notificación automática a responsables.
- Consulta, filtrado y seguimiento del estado de los reportes.

---

## Tipos de Reportes Soportados

El sistema permite la gestión de múltiples tipos de reportes, entre ellos:

- Reportes N2
- Tarjetas de anormalidad
- Reportes de mantenimiento de máquinas
- Ciclos de mejora

Cada tipo de reporte cuenta con su propio flujo de datos, responsables y reglas de notificación, pero todos son gestionados desde una plataforma unificada.

---

## Arquitectura del Sistema

La aplicación está diseñada bajo una **arquitectura de tres capas**, aprovechando el ecosistema de Google Workspace.

### Capa de Presentación (Frontend)

- Desarrollada en **HTML, CSS y JavaScript**
- Formularios dinámicos para el ingreso de información
- Vistas para consulta y seguimiento de reportes
- Comunicación directa con el backend mediante funciones expuestas de Apps Script

### Capa de Lógica de Negocio (Backend)

- Implementada en **Google Apps Script**
- Manejo de validaciones
- Control de flujos de reporte
- Gestión de estados
- Orquestación de envío de correos
- Integración con Google Sheets y Google Drive

### Capa de Persistencia

- **Google Sheets** como base de datos principal
- **Google Drive** para el almacenamiento de imágenes y evidencias
- Relación directa entre registros y archivos almacenados

---

## Tecnologías Utilizadas

| Componente | Tecnología |
|----------|-----------|
| Lenguaje principal | JavaScript |
| Backend | Google Apps Script |
| Frontend | HTML, CSS, JavaScript |
| Base de datos | Google Sheets |
| Almacenamiento de archivos | Google Drive |
| Servicio de correo | Gmail (Apps Script) |
| Despliegue | Apps Script Web App |

---

## Estructura del Proyecto

```text
/
├── Code.js
│   ├── Funciones backend
│   ├── Lógica de negocio
│   ├── Gestión de Sheets y Drive
│   └── Envío de correos
│
├── index.html
│   ├── Interfaz principal
│   ├── Formularios de reporte
│   └── Vistas de consulta
│
├── appsscript.json
│   └── Configuración del proyecto de Apps Script
│
├── assets/
│   └── Recursos estáticos
│
└── README.md
    └── Documentación del proyecto
```

---

## Flujo Funcional del Sistema

### 1. Autenticación de Usuario

El acceso al sistema se realiza mediante la identificación del usuario (cédula).  
La validación se efectúa contra una hoja de cálculo que contiene los usuarios autorizados (líderes y responsables).

---

### 2. Creación de Reportes

El usuario selecciona el tipo de reporte y diligencia el formulario correspondiente, ingresando:

- Información general del reporte
- Descripción del problema o situación
- Área o proceso involucrado
- Evidencias fotográficas (opcional)

---

### 3. Almacenamiento de la Información

- Los datos del reporte se almacenan en **Google Sheets**, en hojas estructuradas según el tipo de reporte.
- Las imágenes se suben automáticamente a **Google Drive** y se vinculan al registro correspondiente.

---

### 4. Notificaciones Automáticas

Una vez registrado el reporte:

- Se envían correos automáticos a los responsables definidos.
- El contenido del correo varía según el tipo de reporte y el estado del flujo.
- Se garantiza trazabilidad y notificación oportuna.

---

### 5. Consulta y Seguimiento

- Los usuarios pueden visualizar los reportes registrados.
- Se permite el filtrado por estado, tipo, fecha y responsable.
- Los administradores pueden gestionar y actualizar el estado de los reportes.

---

## Gestión de Roles

El sistema contempla distintos niveles de acceso:

| Rol | Descripción |
|---|---|
| Usuario | Registro y consulta de reportes propios |
| Administrador | Visualización y gestión global de reportes |
| Sistema | Ejecución automática de procesos y notificaciones |

---

## Google Sheets Utilizados

El sistema utiliza múltiples hojas de cálculo para organizar la información:

- Reportes principales
- Comentarios y seguimientos
- Usuarios y líderes
- Configuración de correos
- Datos auxiliares para formularios dinámicos

Cada hoja cumple una función específica dentro del flujo del sistema.

---

## Configuración del Proyecto

### Requisitos Previos

- Cuenta de Google Workspace
- Acceso a Google Apps Script
- Permisos sobre Google Sheets y Google Drive

---

## Flujo del Sistema

![Arquitectura del sistema](docs/UNDTPST.svg)
[![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/practicantetransfdigital/Unificacion-Data-Pastas)

