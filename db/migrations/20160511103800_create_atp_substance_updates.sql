/*
   miércoles, 11 de mayo de 201610:37:43
   Usuario: istas_SQL
   Servidor: HP-LOLO\SQLEXPRESS
   Base de datos: istas_risctox
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE istas_risctox.dbo.risctox_atp_substance_updates
	(
	id int NOT NULL IDENTITY (1, 1),
	num_atp int NULL,
	num_rd varchar(100) NULL,
	num_cas varchar(100) NULL,
	num_ce_einecs varchar(100) NULL,
	num_ce_elincs varchar(100) NULL,
	identificacion varchar(500) NULL,
	categoria_peligro varchar(8000) NULL,
	clas_fraseh varchar(500) NULL,
	pictograma varchar(500) NULL,
	etiq_fraseh varchar(500) NULL,
	etiq_fraseh_ad varchar(500) NULL,
	limites varchar(8000) NULL,
	notas varchar(50) NULL,
	created_at smalldatetime NULL
	)  ON [PRIMARY]
GO
ALTER TABLE istas_risctox.dbo.risctox_atp_substance_updates ADD CONSTRAINT
	DF_risctox_atp_substance_updates_created_at DEFAULT getDate() FOR created_at
GO
ALTER TABLE istas_risctox.dbo.risctox_atp_substance_updates SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
