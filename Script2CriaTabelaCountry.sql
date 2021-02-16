USE [PaisesInfo]
GO

/****** Object:  Table [dbo].[Country]    Script Date: 15/02/2021 01:11:51 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Country](
	[ISOCodeC] [char](2) NOT NULL,
	[Name] [varchar](100) NOT NULL,
	[CapitalCity] [varchar](100) NOT NULL,
	[PhoneCode] [varchar](3) NOT NULL,
	[ContinentCode] [char](2) NOT NULL,
	[CurrencyISOCode] [varchar](3) NOT NULL,
	[CountryFlag] [varchar](100) NOT NULL,
 CONSTRAINT [PK_Country] PRIMARY KEY CLUSTERED 
(
	[ISOCodeC] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO


