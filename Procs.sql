if exists (select * from sysobjects where id = object_id(N'[dbo].[prc_del_Customers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[prc_del_Customers]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[prc_ins_Customers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[prc_ins_Customers]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[prc_sel_Customers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[prc_sel_Customers]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[prc_sel_Customers_Output]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[prc_sel_Customers_Output]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[prc_upd_Customers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[prc_upd_Customers]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE prc_del_Customers

/* ------------------------------------------------------------
   PROCEDURE:    prc_del_Customers                                      
   
   DESCRIPTION:  Deletes a record from table 'Customers'                                    
   
   AUTHOR:       Brian Lockwood 5/17/00 8:27:19 AM                                  
   ------------------------------------------------------------ */

	@CustomerID                        nchar(10)

	AS DELETE FROM [Customers]

	WHERE 

		[CustomerID]                       = @CustomerID

	RETURN @@ERROR


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE prc_ins_Customers

/* ------------------------------------------------------------
   PROCEDURE:    prc_ins_Customers                                      
   
   DESCRIPTION:  Inserts a record into table 'Customers'                                    
   
   AUTHOR:       Brian Lockwood 5/17/00 8:27:19 AM                                  
   ------------------------------------------------------------ */

	@CustomerID                        nchar(10),
	@CompanyName                       nvarchar(80),
	@ContactName                       nvarchar(60) = NULL,
	@ContactTitle                      nvarchar(60) = NULL,
	@Address                           nvarchar(120) = NULL,
	@City                              nvarchar(30) = NULL,
	@Region                            nvarchar(30) = NULL,
	@PostalCode                        nvarchar(20) = NULL,
	@Country                           nvarchar(30) = NULL,
	@Phone                             nvarchar(48) = NULL,
	@Fax                               nvarchar(48) = NULL

	AS INSERT INTO [Customers]

	(
		[CustomerID],
		[CompanyName],
		[ContactName],
		[ContactTitle],
		[Address],
		[City],
		[Region],
		[PostalCode],
		[Country],
		[Phone],
		[Fax]
	)

	VALUES

	(
		@CustomerID,
		@CompanyName,
		@ContactName,
		@ContactTitle,
		@Address,
		@City,
		@Region,
		@PostalCode,
		@Country,
		@Phone,
		@Fax
	)

	RETURN @@ERROR


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE prc_sel_Customers

/* ------------------------------------------------------------
   PROCEDURE:    prc_sel_Customers                                      
   
   DESCRIPTION:  Selects a record from table 'Customers'                                    
   
   AUTHOR:       Brian Lockwood 5/17/00 8:34:12 AM                                  
   ------------------------------------------------------------ */

	@CustomerID                        nchar(10)

	AS SELECT 

		 [CustomerID],
		 [CompanyName],
		 [ContactName],
		 [ContactTitle],
		 [Address],
		 [City],
		 [Region],
		 [PostalCode],
		 [Country],
		 [Phone],
		 [Fax]

 FROM [Customers]

	WHERE 

		[CustomerID]                       = @CustomerID

	RETURN @@ERROR


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE prc_sel_Customers_Output

/* ------------------------------------------------------------
   PROCEDURE:    prc_sel_Customers_Output                                      
   
   DESCRIPTION:  Selects a record from table 'Customers'                                    
   
   AUTHOR:       Brian Lockwood 3/19/00 11:39:49 AM                                 
   ------------------------------------------------------------ */

	@CustomerID                        nchar(10) OUTPUT,
	@CompanyName                       nvarchar(80) OUTPUT,
	@ContactName                       nvarchar(60) OUTPUT,
	@ContactTitle                      nvarchar(60) OUTPUT,
	@Address                           nvarchar(120) OUTPUT,
	@City                              nvarchar(30) OUTPUT,
	@Region                            nvarchar(30) OUTPUT,
	@PostalCode                        nvarchar(20) OUTPUT,
	@Country                           nvarchar(30) OUTPUT,
	@Phone                             nvarchar(48) OUTPUT,
	@Fax                               nvarchar(48) OUTPUT

	AS SELECT 

		@CustomerID                        = [CustomerID],
		@CompanyName                       = [CompanyName],
		@ContactName                       = [ContactName],
		@ContactTitle                      = [ContactTitle],
		@Address                           = [Address],
		@City                              = [City],
		@Region                            = [Region],
		@PostalCode                        = [PostalCode],
		@Country                           = [Country],
		@Phone                             = [Phone],
		@Fax                               = [Fax]

 FROM [Customers]

	WHERE 

		[CustomerID]                       = @CustomerID

	RETURN @@ERROR


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE prc_upd_Customers

/* ------------------------------------------------------------
   PROCEDURE:    prc_upd_Customers                                      
   
   DESCRIPTION:  Updates a record in table 'Customers'                                      
   
   AUTHOR:       Brian Lockwood 5/17/00 8:27:19 AM                                  
   ------------------------------------------------------------ */

	@CustomerID                        nchar(10),
	@CompanyName                       nvarchar(80),
	@ContactName                       nvarchar(60)  = NULL,
	@ContactTitle                      nvarchar(60)  = NULL,
	@Address                           nvarchar(120)  = NULL,
	@City                              nvarchar(30)  = NULL,
	@Region                            nvarchar(30)  = NULL,
	@PostalCode                        nvarchar(20)  = NULL,
	@Country                           nvarchar(30)  = NULL,
	@Phone                             nvarchar(48)  = NULL,
	@Fax                               nvarchar(48)  = NULL

	AS UPDATE [Customers]

	SET 

		[CustomerID]                       = @CustomerID,
		[CompanyName]                      = @CompanyName,
		[ContactName]                      = @ContactName,
		[ContactTitle]                     = @ContactTitle,
		[Address]                          = @Address,
		[City]                             = @City,
		[Region]                           = @Region,
		[PostalCode]                       = @PostalCode,
		[Country]                          = @Country,
		[Phone]                            = @Phone,
		[Fax]                              = @Fax

	WHERE 

		[CustomerID]                       = @CustomerID

	RETURN @@ERROR


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

