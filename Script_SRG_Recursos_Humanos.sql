USE [master]
GO
/****** Object:  Database [SRG_Recursos_Humanos]    Script Date: 05/30/2014 18:00:26 ******/
CREATE DATABASE [SRG_Recursos_Humanos] ON  PRIMARY 
( NAME = N'Natural_Health_Data', FILENAME = N'C:\Program Files (x86)\Microsoft SQL Server\MSSQL10_50.SQL2008\MSSQL\DATA\SRG_Recursos_Humanos.mdf' , SIZE = 332160KB , MAXSIZE = UNLIMITED, FILEGROWTH = 10%)
 LOG ON 
( NAME = N'Natural_Health_Log', FILENAME = N'C:\Program Files (x86)\Microsoft SQL Server\MSSQL10_50.SQL2008\MSSQL\DATA\SRG_Recursos_Humanos.ldf' , SIZE = 8112KB , MAXSIZE = UNLIMITED, FILEGROWTH = 10%)
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET COMPATIBILITY_LEVEL = 80
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [SRG_Recursos_Humanos].[dbo].[sp_fulltext_database] @action = 'disable'
end
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET ANSI_NULL_DEFAULT OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET ANSI_NULLS OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET ANSI_PADDING OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET ANSI_WARNINGS OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET ARITHABORT OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET AUTO_CLOSE ON
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET AUTO_CREATE_STATISTICS ON
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET AUTO_SHRINK ON
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET AUTO_UPDATE_STATISTICS ON
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET CURSOR_CLOSE_ON_COMMIT OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET CURSOR_DEFAULT  GLOBAL
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET CONCAT_NULL_YIELDS_NULL OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET NUMERIC_ROUNDABORT OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET QUOTED_IDENTIFIER OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET RECURSIVE_TRIGGERS OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET  DISABLE_BROKER
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET AUTO_UPDATE_STATISTICS_ASYNC OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET DATE_CORRELATION_OPTIMIZATION OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET TRUSTWORTHY OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET ALLOW_SNAPSHOT_ISOLATION OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET PARAMETERIZATION SIMPLE
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET READ_COMMITTED_SNAPSHOT OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET HONOR_BROKER_PRIORITY OFF
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET  READ_WRITE
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET RECOVERY SIMPLE
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET  MULTI_USER
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET PAGE_VERIFY TORN_PAGE_DETECTION
GO
ALTER DATABASE [SRG_Recursos_Humanos] SET DB_CHAINING OFF
GO
USE [SRG_Recursos_Humanos]
GO
/****** Object:  Table [dbo].[Cat_Dias_No_Laborales]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Dias_No_Laborales](
	[Dia_No_Laboral_ID] [char](5) NOT NULL,
	[Fecha] [datetime] NULL,
	[Comentarios] [varchar](255) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Dias_No_Laborales] PRIMARY KEY CLUSTERED 
(
	[Dia_No_Laboral_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Departamentos]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Departamentos](
	[Departamento_ID] [char](5) NOT NULL,
	[Nombre] [varchar](50) NULL,
	[Clave] [varchar](10) NULL,
	[Comentarios] [varchar](200) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Departamentos] PRIMARY KEY CLUSTERED 
(
	[Departamento_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Empresas]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Empresas](
	[Empresa_ID] [char](5) NOT NULL,
	[Acronimo] [varchar](20) NULL,
	[Nombre] [varchar](100) NULL,
	[RFC] [varchar](20) NULL,
	[Direccion] [varchar](50) NULL,
	[Colonia] [varchar](50) NULL,
	[Ciudad] [varchar](50) NULL,
	[Estado] [varchar](50) NULL,
	[Codigo_Postal] [varchar](20) NULL,
	[Telefono] [varchar](50) NULL,
	[NOI_ID] [int] NULL,
	[Ruta_COI] [varchar](255) NULL,
	[Ruta_NOI] [varchar](255) NULL,
	[Comentarios] [varchar](255) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Tipo_Nomina] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Empresas] PRIMARY KEY CLUSTERED 
(
	[Empresa_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Roles]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Roles](
	[Rol_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Comentarios] [varchar](255) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Roles] PRIMARY KEY CLUSTERED 
(
	[Rol_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Puestos]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Puestos](
	[Puesto_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Abreviatura] [varchar](10) NULL,
	[Descripcion] [varchar](250) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Puestos] PRIMARY KEY CLUSTERED 
(
	[Puesto_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Parametros]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Parametros](
	[Aplica_Retardos] [char](1) NULL,
	[Tolerancia_Retardos] [numeric](18, 0) NULL,
	[Calcula_Horas_Extra] [char](1) NULL,
	[Horas_Maximas_Turno] [numeric](18, 0) NULL,
	[Comidas_Diarias] [numeric](18, 0) NULL,
	[Ruta_Fotos] [varchar](250) NULL,
	[Impresora_Comidas] [varchar](100) NULL,
	[Dias_Caducidad_Contraseña] [int] NULL,
	[Longitud_Minima_Password] [int] NULL,
	[Intentos_Sesion_Fallidos] [int] NULL,
	[Historico_Password] [int] NULL,
	[Edad_Minima_Contratacion] [int] NULL,
	[Horas_Dobles] [int] NULL,
	[Horas_Triples] [int] NULL,
	[Dias_Falta] [int] NULL,
	[Periodo_Retardos_Dias] [int] NULL,
	[Minutos_Tolerancia] [int] NULL,
	[Email_Sistema] [varchar](200) NULL,
	[Email_Administrador] [varchar](200) NULL,
	[Email_Notificacion] [varchar](200) NULL,
	[Email_validacion] [varchar](200) NULL,
	[Hora_Importacion] [datetime] NULL,
	[Hora_Importacion_Dia] [datetime] NULL,
	[Servidor_SMTP] [varchar](100) NULL,
	[Puerto_SMTP] [int] NULL,
	[PDF_Enfermedad_General] [varchar](10) NULL,
	[PDF_Maternidad] [varchar](10) NULL,
	[PDF_Riesgo_Trabajo] [varchar](10) NULL,
	[PDF_Vacaciones] [varchar](10) NULL,
	[PDF_Alumbramiento] [varchar](10) NULL,
	[PDF_Defuncion] [varchar](10) NULL,
	[PDF_Matrimonio] [varchar](10) NULL,
	[PDF_Falta_Justificada] [varchar](10) NULL,
	[PDF_Permiso_Temporal] [varchar](10) NULL,
	[PDF_Horas_Dobles] [varchar](10) NULL,
	[PDF_Horas_Triples] [varchar](10) NULL,
	[PDF_Falta_InJustificada] [varchar](10) NULL,
	[PDF_Permiso_Goce] [varchar](10) NULL,
	[PDF_Permiso_Sin_Goce] [varchar](10) NULL,
	[PDF_Sancion] [varchar](10) NULL,
	[Aviso_Contratacion] [varchar](10) NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Costo_Comida_Empresa] [money] NULL,
	[Costo_Comida_Empleado] [money] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Nivel_Estudio]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Nivel_Estudio](
	[Nivel_Estudio_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Descripcion] [varchar](255) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Nivel_Estudio] PRIMARY KEY CLUSTERED 
(
	[Nivel_Estudio_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Motivos_Baja]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Motivos_Baja](
	[Motivo_Baja_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Descripcion] [varchar](255) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Motivos_Baja] PRIMARY KEY CLUSTERED 
(
	[Motivo_Baja_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Adm_Bitacora_Importacion]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Bitacora_Importacion](
	[Consecutivo] [decimal](18, 0) IDENTITY(1,1) NOT NULL,
	[Fecha] [datetime] NULL,
	[Comentarios] [varchar](1000) NULL,
	[Hora_Ejecucion] [datetime] NULL,
	[Enviado] [char](2) NULL,
	[Tipo_Importacion] [varchar](50) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Adm_Bitacora_Importacion] PRIMARY KEY CLUSTERED 
(
	[Consecutivo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Adm_Asistencias_Registro_Checadores]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Asistencias_Registro_Checadores](
	[No_Movimiento] [decimal](18, 0) IDENTITY(1,1) NOT NULL,
	[Equipo_ID] [char](5) NULL,
	[Empresa_ID] [char](5) NULL,
	[No_Tarjeta] [numeric](18, 0) NULL,
	[Fecha] [datetime] NULL,
	[Hora] [datetime] NULL,
	[Fecha_Importacion] [datetime] NULL,
	[No_Equipo] [int] NULL,
	[E_S] [char](1) NULL,
	[IP] [varchar](50) NULL,
	[Verificacion] [char](1) NULL,
 CONSTRAINT [PK_Adm_Asistencias_Registro_Checadores] PRIMARY KEY CLUSTERED 
(
	[No_Movimiento] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
CREATE NONCLUSTERED INDEX [IX_Adm_Asistencias_Registro_Checadores_Fecha] ON [dbo].[Adm_Asistencias_Registro_Checadores] 
(
	[Fecha] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
CREATE NONCLUSTERED INDEX [IX_Adm_Asistencias_Registro_Checadores_No_Equipo] ON [dbo].[Adm_Asistencias_Registro_Checadores] 
(
	[No_Equipo] ASC,
	[Equipo_ID] ASC,
	[Empresa_ID] ASC,
	[E_S] ASC,
	[Verificacion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
CREATE NONCLUSTERED INDEX [IX_Fecha] ON [dbo].[Adm_Asistencias_Registro_Checadores] 
(
	[Fecha] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
CREATE NONCLUSTERED INDEX [IX_No_Tarjeta] ON [dbo].[Adm_Asistencias_Registro_Checadores] 
(
	[No_Tarjeta] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Adm_Entradas_Comedor]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Entradas_Comedor](
	[No_Movimiento] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[Empleado_ID] [char](5) NULL,
	[Fecha] [datetime] NULL,
	[Hora] [datetime] NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Adm_Entradas_Comedor] PRIMARY KEY CLUSTERED 
(
	[No_Movimiento] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Cursos]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Cursos](
	[Curso_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Horas] [numeric](18, 2) NULL,
	[Tipo] [varchar](20) NULL,
	[Instructor] [varchar](100) NULL,
	[Comentarios] [varchar](250) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Cursos] PRIMARY KEY CLUSTERED 
(
	[Curso_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cfg_Formatos_Detalles]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cfg_Formatos_Detalles](
	[Nombre] [varchar](50) NOT NULL,
	[Campo] [varchar](100) NULL,
	[X] [float] NULL,
	[Y] [float] NULL,
	[Longitud] [float] NULL,
	[Tipo] [varchar](20) NULL,
	[Formato] [varchar](20) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cfg_Formatos]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cfg_Formatos](
	[Nombre] [varchar](50) NOT NULL,
	[Letra_Generales] [varchar](50) NULL,
	[Estilo_Generales] [varchar](20) NULL,
	[Tamaño_Generales] [int] NULL,
	[No_Columnas] [int] NULL,
	[No_Detalles] [int] NULL,
	[Separacion_Detalles] [float] NULL,
	[Letra_Detalles] [varchar](20) NULL,
	[Estilo_Detalles] [varchar](20) NULL,
	[Tamaño_Detalles] [float] NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Zonas]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Zonas](
	[Zona_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Comentarios] [varchar](250) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Zonas] PRIMARY KEY CLUSTERED 
(
	[Zona_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Tipos_Faltas]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Tipos_Faltas](
	[Tipo_Falta_ID] [char](5) NOT NULL,
	[Descripcion] [varchar](50) NULL,
	[Simbologia] [varchar](10) NULL,
	[Codigo_NOI] [varchar](50) NULL,
	[Comentarios] [varchar](255) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
	[Clasificacion] [varchar](20) NULL,
 CONSTRAINT [PK_Cat_Tipos_Faltas] PRIMARY KEY CLUSTERED 
(
	[Tipo_Falta_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Gaps]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Gaps](
	[Gap_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Comentarios] [varchar](250) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Gaps] PRIMARY KEY CLUSTERED 
(
	[Gap_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Equipos_Identificadores]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Equipos_Identificadores](
	[Equipo_ID] [char](5) NOT NULL,
	[No_Equipo] [int] NOT NULL,
	[Direccion_IP] [varchar](20) NULL,
	[Puerto_IP] [int] NOT NULL,
	[Descripcion] [varchar](100) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Equipos_Identificadores] PRIMARY KEY CLUSTERED 
(
	[Equipo_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Turnos]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Turnos](
	[Turno_ID] [char](5) NOT NULL,
	[Nombre] [varchar](50) NULL,
	[Hora_Inicio] [datetime] NULL,
	[Hora_Termino] [datetime] NULL,
	[Comida_Inicio] [datetime] NULL,
	[Comida_Termino] [datetime] NULL,
	[Horas_Turno] [numeric](18, 2) NULL,
	[Horas_Comida] [numeric](18, 2) NULL,
	[Comentarios] [varchar](255) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Turnos] PRIMARY KEY CLUSTERED 
(
	[Turno_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Tmp_Empleados_Faltas]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Tmp_Empleados_Faltas](
	[Fecha] [varchar](50) NULL,
	[Turno] [varchar](50) NULL,
	[No_Tarjeta] [numeric](18, 0) NULL,
	[Ruta_Imagen] [varchar](250) NULL,
	[Clase] [varchar](100) NULL,
	[Nombre] [varchar](100) NULL,
	[Gerencia] [varchar](50) NULL,
	[Area] [varchar](100) NULL,
	[Supervisor] [varchar](100) NULL,
	[Antiguedad] [varchar](50) NULL,
	[Consecutivo] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[Departamento] [varchar](100) NULL,
 CONSTRAINT [PK_Tmp_Empleados_Faltas] PRIMARY KEY CLUSTERED 
(
	[Consecutivo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Tmp_Empleados_Checadas]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Tmp_Empleados_Checadas](
	[Nombre] [varchar](250) NULL,
	[Empresa] [varchar](100) NULL,
	[Departamento] [varchar](100) NULL,
	[Puesto] [varchar](100) NULL,
	[No_Tarjeta] [int] NULL,
	[Turno] [varchar](100) NULL,
	[Imagen_Perfil] [varchar](250) NULL,
	[Consecutivo] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_Tmp_Empleados_Checadas_1] PRIMARY KEY CLUSTERED 
(
	[Consecutivo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Seguridad_Sistema]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Seguridad_Sistema](
	[Rol_ID] [char](5) NULL,
	[Menu_Habilitado] [varchar](100) NULL,
	[Nombre_Sistema] [varchar](100) NULL,
	[Tipo] [char](10) NULL,
	[Habilitar] [char](1) NULL,
	[Alta] [char](1) NULL,
	[Cambio] [char](1) NULL,
	[Eliminar] [char](1) NULL,
	[Consultar] [char](1) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Transportes]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Transportes](
	[Transporte_ID] [char](5) NOT NULL,
	[Zona_ID] [char](5) NULL,
	[Nombre] [varchar](100) NULL,
	[Comentarios] [varchar](250) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Transportes] PRIMARY KEY CLUSTERED 
(
	[Transporte_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Empresas_Equipos_Identificacion]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Empresas_Equipos_Identificacion](
	[Equipo_ID] [char](5) NULL,
	[Empresa_ID] [char](5) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Usuarios]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Usuarios](
	[Usuario_ID] [char](5) NOT NULL,
	[Rol_ID] [char](5) NULL,
	[Almacen_ID] [char](5) NULL,
	[Estatus] [varchar](20) NULL,
	[Nombre] [varchar](100) NULL,
	[Login] [varchar](20) NULL,
	[Contraseña] [varchar](20) NULL,
	[Fecha_Caduca] [datetime] NULL,
	[Fecha_Ultimo_Cambio_Password] [datetime] NULL,
	[Sesion_Abierta] [varchar](2) NULL,
	[Nombre_Equipo] [varchar](250) NULL,
	[Comentarios] [varchar](250) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[No_Nomina] [int] NULL,
 CONSTRAINT [PK_Cat_Usuarios_1] PRIMARY KEY CLUSTERED 
(
	[Usuario_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Turnos_Detalles]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Turnos_Detalles](
	[Turno_ID] [char](5) NULL,
	[Dia_Semana] [varchar](20) NULL,
	[Hora_Inicio] [datetime] NULL,
	[Hora_Termino] [datetime] NULL,
	[Comida_Inicio] [datetime] NULL,
	[Comida_Termino] [datetime] NULL,
	[Horas_Turno] [numeric](18, 2) NULL,
	[Horas_Comida] [numeric](18, 2) NULL,
	[Dia_Descanso] [char](2) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Empleados](
	[Empleado_ID] [char](5) NOT NULL,
	[Empresa_ID] [char](5) NULL,
	[Supervisor_ID] [char](5) NULL,
	[Departamento_ID] [char](5) NULL,
	[Puesto_ID] [char](5) NULL,
	[Turno_ID] [char](5) NULL,
	[Gap_ID] [char](5) NULL,
	[Nivel_Academico_ID] [char](5) NULL,
	[Motivo_Baja_ID] [char](5) NULL,
	[Transporte_ID] [char](5) NULL,
	[Nomipaq_ID] [varchar](10) NULL,
	[No_Tarjeta] [numeric](18, 0) NULL,
	[Estatus] [char](1) NULL,
	[Tipo] [char](1) NULL,
	[Nombre] [varchar](50) NULL,
	[Apellido_Paterno] [varchar](50) NULL,
	[Apellido_Materno] [varchar](50) NULL,
	[Lugar_Nacimiento] [varchar](100) NULL,
	[Direccion] [varchar](1000) NULL,
	[Colonia] [varchar](100) NULL,
	[Estado] [varchar](100) NULL,
	[Ciudad] [varchar](100) NULL,
	[Codigo_Postal] [varchar](10) NULL,
	[Sexo] [varchar](20) NULL,
	[Fecha_Nacimiento] [datetime] NULL,
	[Estado_Civil] [varchar](50) NULL,
	[Cedula_Identidad_Ciudadana] [varchar](50) NULL,
	[Clave_Elector] [varchar](30) NULL,
	[RFC] [varchar](20) NULL,
	[Curp] [varchar](30) NULL,
	[Nss] [varchar](30) NULL,
	[Imagen_Perfil] [varchar](100) NULL,
	[Fecha_Ingreso] [datetime] NULL,
	[Tipo_Empleado] [varchar](50) NULL,
	[Tipo_Contratacion] [varchar](50) NULL,
	[Fecha_Termino_Contrato] [datetime] NULL,
	[Salario_Diario] [money] NULL,
	[Salario_Diario_Variable] [money] NULL,
	[Trabaja_Domingos] [char](1) NULL,
	[Infonavit] [char](1) NULL,
	[Retardos] [int] NULL,
	[Fecha_Retardo] [datetime] NULL,
	[En_Caso_Emergencia] [varchar](100) NULL,
	[Telefono_Emergencia1] [varchar](100) NULL,
	[Telefono_Emergencia2] [varchar](100) NULL,
	[Alergia1] [varchar](100) NULL,
	[Alergia2] [varchar](100) NULL,
	[Alergia3] [varchar](100) NULL,
	[Gerencia_UAP] [char](5) NULL,
	[Clave_SAP] [varchar](50) NULL,
	[Campo_1] [varchar](500) NULL,
	[Campo_2] [varchar](500) NULL,
	[Campo_3] [varchar](500) NULL,
	[Campo_4] [varchar](500) NULL,
	[Campo_5] [varchar](500) NULL,
	[Fecha_Baja] [datetime] NULL,
	[Comentarios_Baja] [varchar](1000) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Empleados] PRIMARY KEY CLUSTERED 
(
	[Empleado_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Usuarios_Password]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Usuarios_Password](
	[Usuario_ID] [char](5) NULL,
	[Password] [varchar](255) NULL,
	[Fecha_Password] [datetime] NULL,
	[No_Partida] [numeric](18, 0) IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Secciones]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Secciones](
	[Seccion_ID] [char](5) NOT NULL,
	[Supervisor_ID] [char](5) NULL,
	[Clave] [varchar](20) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Secciones] PRIMARY KEY CLUSTERED 
(
	[Seccion_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Gerencias]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Gerencias](
	[Gerencia_ID] [char](5) NOT NULL,
	[Supervisor_ID] [char](5) NULL,
	[Nombre] [varchar](20) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Clave_SAP] [varchar](50) NULL,
 CONSTRAINT [PK_Cat_Gerencias] PRIMARY KEY CLUSTERED 
(
	[Gerencia_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Empleados_Parentesco]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Empleados_Parentesco](
	[Empleado_ID] [char](5) NOT NULL,
	[Nombre] [varchar](100) NULL,
	[Parentesco] [varchar](50) NULL,
	[Fecha_Nacimiento] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Empleados_Huellas]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Empleados_Huellas](
	[Empleado_ID] [char](5) NULL,
	[No_Tarjeta] [numeric](18, 0) NULL,
	[Huella_Digital] [varbinary](max) NULL,
	[Huella_Ruta] [varchar](250) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Empleados_Evaluaciones]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Empleados_Evaluaciones](
	[Consecutivo] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[Empleado_ID] [char](5) NULL,
	[Evaluacion] [varchar](100) NULL,
	[Fecha] [datetime] NULL,
	[Proxima_Evaluacion] [datetime] NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Cat_Empleados_Evaluaciones] PRIMARY KEY CLUSTERED 
(
	[Consecutivo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Cat_Cursos_Detalles]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Cat_Cursos_Detalles](
	[Curso_ID] [char](5) NULL,
	[Empleado_ID] [char](5) NULL,
	[Estatus] [varchar](20) NULL,
	[Comentarios] [varchar](250) NULL,
	[Fecha_Inicio] [datetime] NULL,
	[Fecha_Fin] [datetime] NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](50) NULL,
	[Fecha_Modifico] [datetime] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Adm_Movimientos_Asistencias]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Movimientos_Asistencias](
	[No_Movimiento] [char](10) NOT NULL,
	[Tipo_Incidencia] [char](1) NOT NULL,
	[Empleado_ID] [char](5) NULL,
	[Empresa_ID] [char](5) NULL,
	[Departamento_ID] [char](5) NULL,
	[Tipo_Falta_ID] [char](5) NULL,
	[Fecha_Solicitud] [datetime] NULL,
	[Fecha_Inicio] [datetime] NULL,
	[Fecha_Termino] [datetime] NULL,
	[Dias_Permiso] [int] NULL,
	[Periodo] [int] NULL,
	[Horas_Acuerdo] [decimal](18, 4) NULL,
	[No_Faltas] [int] NULL,
	[Hora_Regreso] [datetime] NULL,
	[Motivo] [varchar](100) NULL,
	[Observaciones] [varchar](255) NULL,
	[Simbologia] [varchar](10) NULL,
	[SubSimbologia] [char](2) NULL,
	[Estatus] [char](1) NULL,
	[Usuario_Creo] [varchar](100) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](100) NULL,
	[Fecha_Modifico] [datetime] NULL,
 CONSTRAINT [PK_Adm_Movimientos_Asistencias] PRIMARY KEY CLUSTERED 
(
	[No_Movimiento] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Adm_Cambios_Turnos]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Cambios_Turnos](
	[Consecutivo] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[Empleado_ID] [char](5) NULL,
	[Turno_Anterior_ID] [char](5) NULL,
	[Turno_Nuevo_ID] [char](5) NULL,
	[Fecha_Cambio] [datetime] NULL,
	[Estatus] [varchar](20) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
 CONSTRAINT [PK_Adm_Cambios_Turnos] PRIMARY KEY CLUSTERED 
(
	[Consecutivo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[Adm_Asistencias_Detalles]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Asistencias_Detalles](
	[No_Operacion] [decimal](18, 0) IDENTITY(1,1) NOT NULL,
	[Empleado_ID] [char](5) NULL,
	[No_Tarjeta] [varchar](10) NULL,
	[Fecha] [datetime] NULL,
	[Hora_Entrada] [datetime] NULL,
	[Hora_Salida] [datetime] NULL,
	[Hora_Comida_Entrada] [datetime] NULL,
	[Hora_Comida_Salida] [datetime] NULL,
	[Horas_Laboradas] [decimal](18, 2) NULL,
	[Validada] [char](1) NULL,
	[Proceso] [varchar](20) NULL,
	[Fecha_Importacion] [datetime] NULL,
	[Fecha_Valido] [datetime] NULL,
	[Horas_Extra] [decimal](18, 2) NULL,
 CONSTRAINT [PK_Adm_Asistencias_Detalles] PRIMARY KEY CLUSTERED 
(
	[No_Operacion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
CREATE NONCLUSTERED INDEX [IX_Fecha] ON [dbo].[Adm_Asistencias_Detalles] 
(
	[Fecha] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
CREATE NONCLUSTERED INDEX [IX_No_Tarjeta] ON [dbo].[Adm_Asistencias_Detalles] 
(
	[No_Tarjeta] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Adm_Asistencias]    Script Date: 05/30/2014 18:00:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[Adm_Asistencias](
	[No_Asistencia] [decimal](18, 0) IDENTITY(1,1) NOT NULL,
	[Supervisor_ID] [char](5) NULL,
	[Empleado_ID] [char](5) NULL,
	[Turno_ID] [char](5) NULL,
	[No_Tarjeta] [varchar](50) NULL,
	[Fecha] [datetime] NULL,
	[Hora_Entrada_Turno] [datetime] NULL,
	[Hora_Salida_Turno] [datetime] NULL,
	[Hora_Entrada_Comida_Turno] [datetime] NULL,
	[Hora_Salida_Comida_Turno] [datetime] NULL,
	[Hora_Entrada_Comida] [datetime] NULL,
	[Hora_Salida_Comida] [datetime] NULL,
	[Hora_Entrada] [datetime] NULL,
	[Hora_Salida] [datetime] NULL,
	[Horas_Extra] [decimal](18, 2) NULL,
	[Horas_Aprobadas] [decimal](18, 2) NULL,
	[Tiempo_Retardo] [int] NULL,
	[Simbologia] [varchar](10) NULL,
	[SubSimbologia] [char](2) NULL,
	[Referencia] [char](10) NULL,
	[Tipo_Incidencia] [char](1) NULL,
	[Usuario_Creo] [varchar](50) NULL,
	[Fecha_Creo] [datetime] NULL,
	[Usuario_Modifico] [varchar](10) NULL,
	[Fecha_Modifico] [datetime] NULL,
	[Horas_Calculadas] [decimal](18, 2) NULL,
 CONSTRAINT [PK_Adm_Asistencias] PRIMARY KEY CLUSTERED 
(
	[No_Asistencia] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
CREATE NONCLUSTERED INDEX [IX_Fecha] ON [dbo].[Adm_Asistencias] 
(
	[Fecha] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
CREATE NONCLUSTERED INDEX [IX_No_Tarjeta] ON [dbo].[Adm_Asistencias] 
(
	[No_Tarjeta] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
GO
/****** Object:  ForeignKey [FK_Seguridad_Sistema_Cat_Roles]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Seguridad_Sistema]  WITH CHECK ADD  CONSTRAINT [FK_Seguridad_Sistema_Cat_Roles] FOREIGN KEY([Rol_ID])
REFERENCES [dbo].[Cat_Roles] ([Rol_ID])
GO
ALTER TABLE [dbo].[Seguridad_Sistema] CHECK CONSTRAINT [FK_Seguridad_Sistema_Cat_Roles]
GO
/****** Object:  ForeignKey [FK_Cat_Transportes_Cat_Zonas]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Transportes]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Transportes_Cat_Zonas] FOREIGN KEY([Zona_ID])
REFERENCES [dbo].[Cat_Zonas] ([Zona_ID])
GO
ALTER TABLE [dbo].[Cat_Transportes] CHECK CONSTRAINT [FK_Cat_Transportes_Cat_Zonas]
GO
/****** Object:  ForeignKey [FK_Cat_Empresas_Equipos_Identificacion_Cat_Empresas]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empresas_Equipos_Identificacion]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empresas_Equipos_Identificacion_Cat_Empresas] FOREIGN KEY([Empresa_ID])
REFERENCES [dbo].[Cat_Empresas] ([Empresa_ID])
GO
ALTER TABLE [dbo].[Cat_Empresas_Equipos_Identificacion] CHECK CONSTRAINT [FK_Cat_Empresas_Equipos_Identificacion_Cat_Empresas]
GO
/****** Object:  ForeignKey [FK_Cat_Empresas_Equipos_Identificacion_Cat_Equipos_Identificadores]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empresas_Equipos_Identificacion]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empresas_Equipos_Identificacion_Cat_Equipos_Identificadores] FOREIGN KEY([Equipo_ID])
REFERENCES [dbo].[Cat_Equipos_Identificadores] ([Equipo_ID])
GO
ALTER TABLE [dbo].[Cat_Empresas_Equipos_Identificacion] CHECK CONSTRAINT [FK_Cat_Empresas_Equipos_Identificacion_Cat_Equipos_Identificadores]
GO
/****** Object:  ForeignKey [FK_Cat_Usuarios_Cat_Roles]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Usuarios]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Usuarios_Cat_Roles] FOREIGN KEY([Rol_ID])
REFERENCES [dbo].[Cat_Roles] ([Rol_ID])
GO
ALTER TABLE [dbo].[Cat_Usuarios] CHECK CONSTRAINT [FK_Cat_Usuarios_Cat_Roles]
GO
/****** Object:  ForeignKey [FK_Cat_Turnos_Detalles_Cat_Turnos]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Turnos_Detalles]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Turnos_Detalles_Cat_Turnos] FOREIGN KEY([Turno_ID])
REFERENCES [dbo].[Cat_Turnos] ([Turno_ID])
GO
ALTER TABLE [dbo].[Cat_Turnos_Detalles] CHECK CONSTRAINT [FK_Cat_Turnos_Detalles_Cat_Turnos]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Departamentos]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Departamentos] FOREIGN KEY([Departamento_ID])
REFERENCES [dbo].[Cat_Departamentos] ([Departamento_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Departamentos]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Empresas]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH NOCHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Empresas] FOREIGN KEY([Empresa_ID])
REFERENCES [dbo].[Cat_Empresas] ([Empresa_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Empresas]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Gaps]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Gaps] FOREIGN KEY([Gap_ID])
REFERENCES [dbo].[Cat_Gaps] ([Gap_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Gaps]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Nivel_Estudio]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Nivel_Estudio] FOREIGN KEY([Nivel_Academico_ID])
REFERENCES [dbo].[Cat_Nivel_Estudio] ([Nivel_Estudio_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Nivel_Estudio]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Puestos]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Puestos] FOREIGN KEY([Puesto_ID])
REFERENCES [dbo].[Cat_Puestos] ([Puesto_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Puestos]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Transportes]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Transportes] FOREIGN KEY([Transporte_ID])
REFERENCES [dbo].[Cat_Transportes] ([Transporte_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Transportes]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Cat_Turnos]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados]  WITH NOCHECK ADD  CONSTRAINT [FK_Cat_Empleados_Cat_Turnos] FOREIGN KEY([Turno_ID])
REFERENCES [dbo].[Cat_Turnos] ([Turno_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados] CHECK CONSTRAINT [FK_Cat_Empleados_Cat_Turnos]
GO
/****** Object:  ForeignKey [FK_Cat_Usuarios_Password_Cat_Usuarios]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Usuarios_Password]  WITH NOCHECK ADD  CONSTRAINT [FK_Cat_Usuarios_Password_Cat_Usuarios] FOREIGN KEY([Usuario_ID])
REFERENCES [dbo].[Cat_Usuarios] ([Usuario_ID])
GO
ALTER TABLE [dbo].[Cat_Usuarios_Password] CHECK CONSTRAINT [FK_Cat_Usuarios_Password_Cat_Usuarios]
GO
/****** Object:  ForeignKey [FK_Cat_Secciones_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Secciones]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Secciones_Cat_Empleados] FOREIGN KEY([Supervisor_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Cat_Secciones] CHECK CONSTRAINT [FK_Cat_Secciones_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Cat_Gerencias_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Gerencias]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Gerencias_Cat_Empleados] FOREIGN KEY([Supervisor_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Cat_Gerencias] CHECK CONSTRAINT [FK_Cat_Gerencias_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Parentesco_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados_Parentesco]  WITH NOCHECK ADD  CONSTRAINT [FK_Cat_Empleados_Parentesco_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados_Parentesco] CHECK CONSTRAINT [FK_Cat_Empleados_Parentesco_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Huellas_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados_Huellas]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Huellas_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados_Huellas] CHECK CONSTRAINT [FK_Cat_Empleados_Huellas_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Cat_Empleados_Evaluaciones_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Empleados_Evaluaciones]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Empleados_Evaluaciones_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Cat_Empleados_Evaluaciones] CHECK CONSTRAINT [FK_Cat_Empleados_Evaluaciones_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Cat_Cursos_Detalles_Cat_Cursos]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Cursos_Detalles]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Cursos_Detalles_Cat_Cursos] FOREIGN KEY([Curso_ID])
REFERENCES [dbo].[Cat_Cursos] ([Curso_ID])
GO
ALTER TABLE [dbo].[Cat_Cursos_Detalles] CHECK CONSTRAINT [FK_Cat_Cursos_Detalles_Cat_Cursos]
GO
/****** Object:  ForeignKey [FK_Cat_Cursos_Detalles_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Cat_Cursos_Detalles]  WITH CHECK ADD  CONSTRAINT [FK_Cat_Cursos_Detalles_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Cat_Cursos_Detalles] CHECK CONSTRAINT [FK_Cat_Cursos_Detalles_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Adm_Movimientos_Asistencias_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Adm_Movimientos_Asistencias]  WITH NOCHECK ADD  CONSTRAINT [FK_Adm_Movimientos_Asistencias_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Adm_Movimientos_Asistencias] CHECK CONSTRAINT [FK_Adm_Movimientos_Asistencias_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Adm_Cambios_Turnos_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Adm_Cambios_Turnos]  WITH CHECK ADD  CONSTRAINT [FK_Adm_Cambios_Turnos_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Adm_Cambios_Turnos] CHECK CONSTRAINT [FK_Adm_Cambios_Turnos_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Adm_Cambios_Turnos_Cat_Turnos]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Adm_Cambios_Turnos]  WITH CHECK ADD  CONSTRAINT [FK_Adm_Cambios_Turnos_Cat_Turnos] FOREIGN KEY([Turno_Nuevo_ID])
REFERENCES [dbo].[Cat_Turnos] ([Turno_ID])
GO
ALTER TABLE [dbo].[Adm_Cambios_Turnos] CHECK CONSTRAINT [FK_Adm_Cambios_Turnos_Cat_Turnos]
GO
/****** Object:  ForeignKey [FK_Adm_Asistencias_Detalles_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Adm_Asistencias_Detalles]  WITH NOCHECK ADD  CONSTRAINT [FK_Adm_Asistencias_Detalles_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Adm_Asistencias_Detalles] CHECK CONSTRAINT [FK_Adm_Asistencias_Detalles_Cat_Empleados]
GO
/****** Object:  ForeignKey [FK_Adm_Asistencias_Cat_Empleados]    Script Date: 05/30/2014 18:00:29 ******/
ALTER TABLE [dbo].[Adm_Asistencias]  WITH NOCHECK ADD  CONSTRAINT [FK_Adm_Asistencias_Cat_Empleados] FOREIGN KEY([Empleado_ID])
REFERENCES [dbo].[Cat_Empleados] ([Empleado_ID])
GO
ALTER TABLE [dbo].[Adm_Asistencias] CHECK CONSTRAINT [FK_Adm_Asistencias_Cat_Empleados]
GO
