USE [PaisesInfo]
GO

/****** Object:  Table [dbo].[CountryLanguage]    Script Date: 15/02/2021 01:41:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CountryLanguage](
	[ISOCodeC] [char](2) NOT NULL,
	[ISOCodeL] [char](3) NOT NULL
) ON [PRIMARY]
GO


