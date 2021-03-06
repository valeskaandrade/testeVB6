/*
   segunda-feira, 15 de fevereiro de 202101:44:43
   Usuário: 
   Servidor: DESKTOP-MP61253
   Banco de Dados: PaisesInfo
   Aplicativo: 
*/

/* Para impedir possíveis problemas de perda de dados, analise este script detalhadamente antes de executá-lo fora do contexto do designer de banco de dados.*/
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
ALTER TABLE dbo.Language SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Country SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.CountryLanguage ADD CONSTRAINT
	PK_CountryLanguage PRIMARY KEY CLUSTERED 
	(
	ISOCodeC,
	ISOCodeL
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.CountryLanguage ADD CONSTRAINT
	FK_CountryLanguage_Country FOREIGN KEY
	(
	ISOCodeC
	) REFERENCES dbo.Country
	(
	ISOCodeC
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
ALTER TABLE dbo.CountryLanguage ADD CONSTRAINT
	FK_CountryLanguage_Language FOREIGN KEY
	(
	ISOCodeL
	) REFERENCES dbo.Language
	(
	ISOCodeL
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
ALTER TABLE dbo.CountryLanguage SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
