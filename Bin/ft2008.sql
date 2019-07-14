--===========================================================================
--����˵����
--1����SQLServer2008���ù���ԱȨ��(����ͬ)ִ�����´��롣
--2������˳����ò�Ҫ�ҡ����빦���ǽ������ݿ⡢�˺š����洢���̵���Ϣ��
--3���½����ݿ���Ϊ��db_FT���½���sysadminȨ�޵��˺ţ�FT_MS������ftms��
--4���������ݿ�IP�����ݿ������˺����������õ�����˳����С�
--===========================================================================

--==================�ָ���===================================================

--===========================================================================
--�������������ݿ�
--===========================================================================
USE [master]
GO

/****** �����ݿ�db_FT�Ѵ��ڣ���ɾ��******/
IF EXISTS (SELECT 1 FROM sys.sysdatabases WHERE name ='db_FT')
	BEGIN
		--�ر��������ݿ�db_FT���������ӡ���������ʱɾ�����˸����ݿ⡣
		DECLARE @spid_db INT

		DECLARE CUR_db CURSOR FOR 
		SELECT spid FROM sys.sysprocesses WHERE dbid = DB_ID('db_FT');

		OPEN CUR_db

		FETCH NEXT FROM CUR_db INTO @spid_db

		WHILE @@FETCH_STATUS = 0
		BEGIN
			EXEC ('KILL ' + @spid_db)
			FETCH NEXT FROM CUR_db INTO @spid_db
		END
		CLOSE CUR_db
		DEALLOCATE CUR_db

		--ɾ��ָ�����ݿ⡣
		DROP DATABASE db_FT
	END
	

/****** Object:  Database [db_FT]    Script Date: 09/18/2018 08:39:13 ******/
CREATE DATABASE [db_FT] ON  PRIMARY 
( NAME = N'db_FT', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL10.MSSQLSERVER\MSSQL\DATA\db_FT.mdf' , SIZE = 3072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'db_FT_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL10.MSSQLSERVER\MSSQL\DATA\db_FT_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO

ALTER DATABASE [db_FT] SET COMPATIBILITY_LEVEL = 100
GO

IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [db_FT].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO

ALTER DATABASE [db_FT] SET ANSI_NULL_DEFAULT OFF 
GO

ALTER DATABASE [db_FT] SET ANSI_NULLS OFF 
GO

ALTER DATABASE [db_FT] SET ANSI_PADDING OFF 
GO

ALTER DATABASE [db_FT] SET ANSI_WARNINGS OFF 
GO

ALTER DATABASE [db_FT] SET ARITHABORT OFF 
GO

ALTER DATABASE [db_FT] SET AUTO_CLOSE OFF 
GO

ALTER DATABASE [db_FT] SET AUTO_CREATE_STATISTICS ON 
GO

ALTER DATABASE [db_FT] SET AUTO_SHRINK OFF 
GO

ALTER DATABASE [db_FT] SET AUTO_UPDATE_STATISTICS ON 
GO

ALTER DATABASE [db_FT] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO

ALTER DATABASE [db_FT] SET CURSOR_DEFAULT  GLOBAL 
GO

ALTER DATABASE [db_FT] SET CONCAT_NULL_YIELDS_NULL OFF 
GO

ALTER DATABASE [db_FT] SET NUMERIC_ROUNDABORT OFF 
GO

ALTER DATABASE [db_FT] SET QUOTED_IDENTIFIER OFF 
GO

ALTER DATABASE [db_FT] SET RECURSIVE_TRIGGERS OFF 
GO

ALTER DATABASE [db_FT] SET  DISABLE_BROKER 
GO

ALTER DATABASE [db_FT] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO

ALTER DATABASE [db_FT] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO

ALTER DATABASE [db_FT] SET TRUSTWORTHY OFF 
GO

ALTER DATABASE [db_FT] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO

ALTER DATABASE [db_FT] SET PARAMETERIZATION SIMPLE 
GO

ALTER DATABASE [db_FT] SET READ_COMMITTED_SNAPSHOT OFF 
GO

ALTER DATABASE [db_FT] SET HONOR_BROKER_PRIORITY OFF 
GO

ALTER DATABASE [db_FT] SET  READ_WRITE 
GO

ALTER DATABASE [db_FT] SET RECOVERY FULL 
GO

ALTER DATABASE [db_FT] SET  MULTI_USER 
GO

ALTER DATABASE [db_FT] SET PAGE_VERIFY CHECKSUM  
GO

ALTER DATABASE [db_FT] SET DB_CHAINING OFF 
GO


--===========================================================================
--����������FTϵͳ��ר���˺�FT_MS������ftms
--===========================================================================
/******���Ѵ����˺�FT_MS������ɾ��֮******/
IF EXISTS (SELECT 1 FROM sys.syslogins WHERE name ='FT_MS')
BEGIN
	--EXEC sp_who 'FT_MS'
	--�Ͽ�ר���˺ŵ��������ӣ����Ͽ�����ɾ����
	DECLARE @spid_login INT

	DECLARE CUR_login CURSOR FOR 
	SELECT spid FROM sys.sysprocesses WHERE loginame ='FT_MS'

	OPEN CUR_login

	FETCH NEXT FROM CUR_login INTO @spid_login

	WHILE @@FETCH_STATUS = 0
	BEGIN
		EXEC ('KILL ' + @spid_login)
		FETCH NEXT FROM CUR_login INTO @spid_login
	END
	CLOSE CUR_login
	DEALLOCATE CUR_login

	--ɾ���Ѵ��ڵ�ר���˺š�
	DROP LOGIN [FT_MS]
END

/****** Object:  Login [FT_MS]    Script Date: 2018/10/11 23:01:54 ******/
CREATE LOGIN [FT_MS] WITH PASSWORD=N'ftms', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[��������], CHECK_EXPIRATION=OFF, CHECK_POLICY=ON
GO

ALTER LOGIN [FT_MS] ENABLE
GO

EXEC sys.sp_addsrvrolemember @loginame = N'FT_MS', @rolename = N'sysadmin'
GO


--==================�ָ���===================================================

--===========================================================================
--������������[tb_FT_Sys_User]�������˺�����
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_User]    Script Date: 2018/9/15 21:48:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tb_FT_Sys_User](
	[UserAutoID] [int] IDENTITY(2000,1) NOT NULL,
	[UserLoginName] [nvarchar](50) NOT NULL,
	[UserPassword] [nvarchar](50) NOT NULL,
	[UserFullName] [nvarchar](50) NULL,
	[UserSex] [nvarchar](2) NULL,
	[UserState] [nvarchar](50) NULL,
	[DeptID] [int] NULL,
	[UserMemo] [nvarchar](500) NULL,
	[FileID] [bigint] NULL,
 CONSTRAINT [PK_tb_FT_Sys_User] PRIMARY KEY CLUSTERED 
(
	[UserAutoID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

--===========================================================================
--��������[tb_FT_Sys_User]���в���Ĭ�ϵĳ���Ȩ���˺�����
--===========================================================================
USE [db_FT]
GO

INSERT INTO [db_FT].[dbo].[tb_FT_Sys_User]([UserLoginName],[UserPassword],[UserFullName])
VALUES('admin','9E7445656E63AB22FC3EA4387D00','��������Ա')	--����ftadmin
GO

INSERT INTO [db_FT].[dbo].[tb_FT_Sys_User]([UserLoginName],[UserPassword],[UserFullName])
VALUES('system','9E7445657C63B622E23EB93876000744','ϵͳ����Ա')	--����ftsystem
GO

--===========================================================================
--������������[tb_FT_Sys_Department]
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_Department]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_FT_Sys_Department](
	[DeptID] [int] IDENTITY(1000,1) NOT NULL,
	[DeptName] [nvarchar](50) NOT NULL,
	[ParentID] [int] NULL,
 CONSTRAINT [PK_tb_FT_Department] PRIMARY KEY CLUSTERED 
(
	[DeptID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

--===========================================================================
--������������[tb_FT_Sys_Func]
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_Func]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_FT_Sys_Func](
	[FuncAutoID] [int] IDENTITY(1000,1) NOT NULL,
	[FuncName] [nvarchar](50) NOT NULL,
	[FuncCaption] [nvarchar](50) NOT NULL,
	[FuncType] [nvarchar](50) NOT NULL,
	[FuncParentID] [int] NOT NULL
) ON [PRIMARY]

GO

--===========================================================================
--���������[tb_FT_Sys_Func]������ϸȨ����Ϣ
--===========================================================================
USE [db_FT]
GO
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('frmSysDepartment' ,'���Ź���' ,'����' ,'1010')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command1' ,'��Ӳ���' ,'��ť' ,'1000')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command2' ,'�޸Ĳ�����Ϣ' ,'��ť' ,'1000')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('frmSysUser' ,'�û�����' ,'����' ,'1010')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command1' ,'����û�' ,'��ť' ,'1003')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command2' ,'�޸��û���Ϣ' ,'��ť' ,'1003')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('TreeView1' ,'�û��б�' ,'����' ,'1003')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('frmSysFunc' ,'��������' ,'����' ,'1010')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('frmSysRole' ,'��ɫ����' ,'����' ,'1010')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('frmSysLog' ,'��־�鿴' ,'����' ,'1010')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Sys' ,'ϵͳ' ,'���˵�' ,'0')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command1' ,'��ӹ���' ,'��ť' ,'1007')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command2' ,'�޸Ĺ�����Ϣ' ,'��ť' ,'1007')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command3' ,'�����ָ����ɫ�������' ,'��ť' ,'1007')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('TreeView1' ,'���ƹ����б�' ,'����' ,'1007')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('TreeView1' ,'�����б�' ,'����' ,'1000')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command3' ,'�û���ɫָ���������' ,'��ť' ,'1003')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command1' ,'��ӽ�ɫ' ,'��ť' ,'1008')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command2' ,'�޸Ľ�ɫ��Ϣ' ,'��ť' ,'1008')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('TreeView1' ,'��ɫ�б�' ,'����' ,'1008')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Combo1' ,'����������ɫȨ��' ,'����' ,'1008')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command3' ,'�����ɫȨ�޷�����' ,'��ť' ,'1008')
INSERT INTO [dbo].[tb_FT_Sys_Func]([FuncName] ,[FuncCaption] ,[FuncType] ,[FuncParentID])
VALUES('Command1' ,'��ѯ' ,'��ť' ,'1009')
GO

-- SQLServer�Զ������Ĳ������ݽű���
-- SET IDENTITY_INSERT [dbo].[tb_FT_Sys_Func] ON
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1000, N'frmSysDepartment', N'���Ź���', N'����', 1010)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1001, N'Command1', N'��Ӳ���', N'��ť', 1000)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1002, N'Command2', N'�޸Ĳ�����Ϣ', N'��ť', 1000)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1003, N'frmSysUser', N'�û�����', N'����', 1010)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1004, N'Command1', N'����û�', N'��ť', 1003)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1005, N'Command2', N'�޸��û���Ϣ', N'��ť', 1003)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1006, N'TreeView1', N'�û��б�', N'����', 1003)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1007, N'frmSysFunc', N'��������', N'����', 1010)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1008, N'frmSysRole', N'��ɫ����', N'����', 1010)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1009, N'frmSysLog', N'��־�鿴', N'����', 1010)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1010, N'Sys', N'ϵͳ', N'���˵�', 0)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1011, N'Command1', N'��ӹ���', N'��ť', 1007)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1012, N'Command2', N'�޸Ĺ�����Ϣ', N'��ť', 1007)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1013, N'Command3', N'�����ָ����ɫ�������', N'��ť', 1007)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1014, N'TreeView1', N'���ƹ����б�', N'����', 1007)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1015, N'TreeView1', N'�����б�', N'����', 1000)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1016, N'Command3', N'�û���ɫָ���������', N'��ť', 1003)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1017, N'Command1', N'��ӽ�ɫ', N'��ť', 1008)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1018, N'Command2', N'�޸Ľ�ɫ��Ϣ', N'��ť', 1008)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1019, N'TreeView1', N'��ɫ�б�', N'����', 1008)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1020, N'Combo1', N'����������ɫȨ��', N'����', 1008)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1021, N'Command3', N'�����ɫȨ�޷�����', N'��ť', 1008)
-- INSERT [dbo].[tb_FT_Sys_Func] ([FuncAutoID], [FuncName], [FuncCaption], [FuncType], [FuncParentID]) VALUES (1022, N'Command1', N'��ѯ', N'��ť', 1009)
-- SET IDENTITY_INSERT [dbo].[tb_FT_Sys_Func] OFF



--===========================================================================
--������������[tb_FT_Sys_OperationLog]
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_OperationLog]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_FT_Sys_OperationLog](
	[LogID] [bigint] IDENTITY(1,1) NOT NULL,
	[LogType] [nvarchar](50) NULL,
	[LogContent] [nvarchar](200) NULL,
	[LogTime] [datetime] NULL,
	[LogTable] [nvarchar](50) NULL,
	[LogFormName] [nvarchar](50) NULL,
	[LogUserFullName] [nvarchar](50) NULL,
	[LogPCIP] [nvarchar](50) NULL,
	[LogPCName] [nvarchar](50) NULL,
 CONSTRAINT [PK_tb_FT_OperationLog] PRIMARY KEY CLUSTERED 
(
	[LogID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

--===========================================================================
--������������[tb_FT_Sys_Role]
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_Role]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_FT_Sys_Role](
	[RoleAutoID] [int] IDENTITY(1000,1) NOT NULL,
	[RoleName] [nvarchar](50) NOT NULL,
	[DeptID] [int] NULL
) ON [PRIMARY]

GO

--===========================================================================
--������������[tb_FT_Sys_RoleFunc]
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_RoleFunc]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_FT_Sys_RoleFunc](
	[RoleAutoID] [int] NOT NULL,
	[FuncAutoID] [int] NOT NULL
) ON [PRIMARY]

GO

--===========================================================================
--������������[tb_FT_Sys_UserRole]
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Sys_UserRole]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_FT_Sys_UserRole](
	[UserAutoID] [int] NOT NULL,
	[RoleAutoID] [int] NOT NULL
) ON [PRIMARY]

GO


--===========================================================================
--������������[tb_FT_Lib_File]�ļ�����
--===========================================================================
USE [db_FT]
GO

/****** Object:  Table [dbo].[tb_FT_Lib_File]    Script Date: 06/06/2019 08:39:43 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tb_FT_Lib_File](
	[FileID] [bigint] IDENTITY(1,1) NOT NULL,
	[FileClassify] [nvarchar](20) NULL,
	[FileExtension] [nvarchar](20) NULL,
	[FileOldName] [nvarchar](50) NULL,
	[FileSaveName] [nvarchar](50) NULL,
	[FileSize] [bigint] NULL,
	[FileSaveLocation] [nvarchar](50) NULL,
	[FileUploadMen] [nvarchar](20) NULL,
	[FileUploadTime] [datetime] NULL
) ON [PRIMARY]

GO


--==================�ָ���===================================================

--===========================================================================
--�����������洢����[sp_FT_Sys_UserLogin]
--===========================================================================
USE [db_FT]
GO

/****** Object:  StoredProcedure [dbo].[sp_FT_Sys_UserLogin]    Script Date: 2018/9/15 23:27:51 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_FT_Sys_UserLogin] 
	-- Add the parameters for the stored procedure here
	@strUN AS NVARCHAR(50)
	,@strPWD AS NVARCHAR(50)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	SELECT * 
	FROM tb_FT_Sys_User 
	WHERE UserLoginName = @strUN AND UserPassword =@strPWD 

END

GO

--===========================================================================
--�����������洢����[sp_FT_Sys_UserInfo]
--===========================================================================
USE [db_FT]
GO
/****** Object:  StoredProcedure [dbo].[sp_FT_Sys_UserInfo]    Script Date: 2018/10/14 16:42:21 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_FT_Sys_UserInfo]
	-- Add the parameters for the stored procedure here
	@intUID AS INT
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	SELECT * FROM tb_FT_Sys_User 
	WHERE UserAutoID = @intUID 
END

GO

--===========================================================================
--�����������洢����[sp_FT_Sys_LogAdd]
--===========================================================================
USE [db_FT]
GO
/****** Object:  StoredProcedure [dbo].[sp_FT_Sys_LogAdd]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_FT_Sys_LogAdd] 
	-- Add the parameters for the stored procedure here
	@strType AS NVARCHAR(50)='select'
	,@strForm AS NVARCHAR(50)=''
	,@strTable AS NVARCHAR(50)=''
	,@strContent AS NVARCHAR(200)=''
	,@strUser AS NVARCHAR(50)=''
	,@strIP AS NVARCHAR(50)=''
	,@strPC AS NVARCHAR(50)=''
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO tb_FT_Sys_OperationLog
	(LogType ,LogFormName ,LogTable ,LogContent ,LogUserFullName ,
	LogPCIP ,LogPCName ,LogTime )
	VALUES(@strType ,@strForm ,@strTable ,@strContent ,@strUser ,
	@strIP ,@strPC ,GETDATE() );
	
END

GO

--===========================================================================
--�����������洢����[sp_FT_Sys_LogQuery]
--===========================================================================
USE [db_FT]
GO
/****** Object:  StoredProcedure [dbo].[sp_FT_Sys_LogQuery]    Script Date: 2018/10/11 22:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_FT_Sys_LogQuery] 
	-- Add the parameters for the stored procedure here
	@strType AS NVARCHAR(50)=''
	,@strContent AS NVARCHAR(200)=''
	,@strTimeA AS NVARCHAR(30)=''
	,@strTimeB AS NVARCHAR(30)=''
	,@strForm AS NVARCHAR(50)=''
	,@strUser AS NVARCHAR(50)=''
	,@strIP AS NVARCHAR(50)=''
	,@strPC AS NVARCHAR(50)=''
	,@strField AS NVARCHAR(50)='LogTime'
	,@strSort AS NVARCHAR(10)='ASC'
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
	
	DECLARE @strSQL AS NVARCHAR(2000) 
	DECLARE @intLoc AS INT
    -- Insert statements for procedure here
    SET @strSQL ='SELECT * FROM tb_FT_Sys_OperationLog '
    
    IF LEN(@strType)>0 SET @strSQL =@strSQL +' AND LogType='''+@strType+'''' 
    IF LEN(@strTimeA)>0 AND LEN(@strTimeB)>0 SET @strSQL =@strSQL +' AND LogTime BETWEEN '''+@strTimeA+''' AND '''+@strTimeB+'''' 
    IF LEN(@strForm)>0 SET @strSQL =@strSQL +' AND LogFormName LIKE ''%'+@strForm+'%''' 
    IF LEN(@strUser)>0 SET @strSQL =@strSQL +' AND LogUserFullName LIKE ''%'+@strUser+'%''' 
    IF LEN(@strIP)>0 SET @strSQL =@strSQL +' AND LogPCIP LIKE ''%'+@strIP+'%''' 
    IF LEN(@strPC)>0 SET @strSQL =@strSQL +' AND LogPCName LIKE ''%'+@strPC+'%'''
    IF LEN(@strContent)>0 SET @strSQL =@strSQL +' AND LogContent LIKE ''%'+@strContent+'%'''
    
    SET @intLoc = CHARINDEX(' AND ',@strSQL)
    IF @intLoc >0 SET @strSQL =STUFF (@strSQL,@intLoc,5,' WHERE ')
    
	SET @strSQL =@strSQL+' ORDER BY '+@strField+' '+@strSort  
  
	EXEC(@strSQL)
	
END

GO

--===========================================================================
--������
--===========================================================================
