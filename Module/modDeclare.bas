Attribute VB_Name = "modDeclare"
Option Explicit


'''��������ָ״
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long  'SetCursorȷ�������״
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
        ByVal lpCursorName As String) As Long   'LoadCursor����ָ�������Դ
Public Const IDC_HAND = "#32649"


'''���Ҵ��ڣ�������Ϣ
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'''ʹ�� ShellExecute ���ļ���ִ�г���
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'hWnd������ָ�������ھ�������������ù��̳��ִ���ʱ��������ΪWindows��Ϣ���ڵĸ�����
'Operation������ָ��Ҫ���еĲ���������:
'''edit �ñ༭���� lpFile ָ�����ĵ������ lpFile �����ĵ������ʧ��;
'''explore ��� lpFile ָ�����ļ���
'''find ���� lpDirectory ָ����Ŀ¼
'''open �� lpFile �ļ���lpFile �������ļ����ļ���
'''print ��ӡ lpFile����� lpFile �����ĵ�������ʧ��
'''properties ��ʾ����
'''runas �����Թ���ԱȨ�����У������Թ���ԱȨ������ĳ��exe
'''NULL ִ��Ĭ�ϡ�open������
'FileName������ָ��Ҫ�򿪵��ļ�����Ҫִ�еĳ����ļ�����Ҫ������ļ�����
'Parameters����FileName������һ����ִ�г�����˲���ָ�������в���������˲���ӦΪnil��PChar(0)
'Directory������ָ��Ĭ��Ŀ¼
'ShowCmd����FileName������һ����ִ�г�����˲���ָ�����򴰿ڵĳ�ʼ��ʾ��ʽ������˲���Ӧ����Ϊ0
'��ShellExecute�������óɹ����򷵻�ֵΪ��ִ�г����ʵ�������������ֵС��32�����ʾ���ִ���,��������:
Public Const NO_ERROR = 0   'ϵͳ�ڴ����Դ����
Public Const ERROR_FILE_NOT_FOUND = 2&  '�Ҳ���ָ�����ļ�
Public Const ERROR_PATH_NOT_FOUND = 3&  '�Ҳ���ָ��·��
Public Const ERROR_BAD_FORMAT = 11&     '.exe�ļ���Ч
Public Const SE_ERR_ACCESSDENIED = 5    '�ܾ�����ָ���ļ�
Public Const SE_ERR_ASSOCINCOMPLETE = 27    '�ļ���������Ч������
Public Const SE_ERR_DDEBUSY = 30    'DDE�������ڴ���DDE�����޷����
Public Const SE_ERR_DDEFAIL = 29    'DDE����ʧ��
Public Const SE_ERR_DDETIMEOUT = 28 '����ʱ���޷����DDE��������
Public Const SE_ERR_DLLNOTFOUND = 32    'δ�ҵ�ָ��dll
Public Const SE_ERR_FNF = 2         'δ�ҵ�ָ���ļ�
Public Const SE_ERR_NOASSOC = 31    'δ�ҵ�������ļ���չ��������Ӧ�ó��򣬱����ӡ���ɴ�ӡ���ļ���
Public Const SE_ERR_OOM = 8         '�ڴ治�㣬�޷���ɲ���
Public Const SE_ERR_PNF = 3         'δ�ҵ�ָ��·��
Public Const SE_ERR_SHARE = 26      '���������ͻ
'ShellExecute����nShowCmd���õĳ���ShowWindow() Commands
Public Const SW_HIDE = 0        '���ش��ڣ��״̬����һ������
Public Const SW_SHOWNORMAL = 1  '��SW_RESTORE��ͬ
Public Const SW_NORMAL = 1      '
Public Const SW_SHOWMINIMIZED = 2   '��С�����ڣ������伤��
Public Const SW_SHOWMAXIMIZED = 3   'SHOWMAXIMIZED ��󻯴��ڣ������伤��
Public Const SW_MAXIMIZE = 3        '
Public Const SW_SHOWNOACTIVATE = 4  '������Ĵ�С��λ����ʾһ�����ڣ�ͬʱ���ı�����
Public Const SW_SHOW = 5            '�õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
Public Const SW_MINIMIZE = 6        '��С�����ڣ��״̬����һ������
Public Const SW_SHOWMINNOACTIVE = 7 '��С��һ�����ڣ�ͬʱ���ı�����
Public Const SW_SHOWNA = 8          '�õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ����ı�����
Public Const SW_RESTORE = 9         '��ԭ���Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
Public Const SW_SHOWDEFAULT = 10    '
Public Const SW_MAX = 10            '

'''ע������API������
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const HKEY_USER_RUN As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"  '��������Զ�����ע����Ӽ�λ��

Public Enum genumRegRootDirectory   'ע������
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Enum genumRegDataType    'ע���ֵ����
    REG_SZ = 1          ' Unicode nul terminated string
    REG_EXPAND_SZ = 2   ' Unicode nul terminated string
    REG_BINARY = 3      ' Free form binary
    REG_DWORD = 4       ' 32-bit number
End Enum

Public Enum genumRegOperateType 'ע����������
    RegRead = 1
    RegWrite = 2
    RegDelete = 3
End Enum

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '������ͣ���У����룩

'���ص�����ϢAPI
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Enum genumComputerInfoType   'Ҫ���صĵ����ϵ���Ϣ���
    ciComputerName
    ciUserName
End Enum


'''����API����Shell_NotifyIcon��һ�ѳ�����ö�١��ṹ�嶼�й�����
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    lpData As gtypeNOTIFYICONDATA) As Long

Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const IMAGE_ICON = 1

Public Const NIF_ICON = &H2     'hIcon��Ա������
Public Const NIF_INFO = &H10    'ʹ��������ʾ ������ͨ����ʾ��
Public Const NIF_MESSAGE = &H1  'uCallbackMessage��Ա������
Public Const NIF_STATE = &H8    'dwState��dwStateMask��Ա������
Public Const NIF_TIP = &H4      'szTip��Ա������

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NIM_VERSION = &H5

Public Const WM_USER As Long = &H400
Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Const NOTIFYICON_VERSION = 3 'ʹ��Windows2000�����һ������ֵ0��ʾʹ��Windows95���

Public Const NIS_HIDDEN = &H1       'ͼ������
Public Const NIS_SHAREDICON = &H2   'ͼ�깲��

Public Const WM_NOTIFY As Long = &H4E
Public Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10

Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209

Public Type gtypeNOTIFYICONDATA
    cbSize As Long  '�ṹ��С���ֽڣ�
    hwnd As Long    '������Ϣ�Ĵ��ھ��
    uID As Long     '����ͼ��ı�ʶ��
    uFlags As Long  '�˳�Ա������Щ������Ա������
    uCallbackMessage As Long        'Ӧ�ó��������Ϣ��ʾ
    hIcon As Long                   '����ͼ����
    szTip As String * 128   '��ʾ��Ϣ,��֪Ϊ�Σ����Ȳ���128����������

    dwState As Long         'ͼ��״̬
    dwStateMask As Long     'ָ��dwState��Ա����Щλ���Ա����û����
    szInfo As String * 256      '������ʾ��Ϣ
    uTimeoutOrVersion As Long  '������ʾ��ʧʱ���汾
    szInfoTitle As String * 64  '������ʾ����
    dwInfoFlags As Long         '��������ʾ������һ��ͼ��
End Type

Public Enum genumNotifyIconFlag
    NIIF_NONE = &H0     'û��ͼ��
    NIIF_INFO = &H1     '��Ϣͼ��
    NIIF_WARNING = &H2  '����ͼ��
    NIIF_ERROR = &H3    '����ͼ��
    NIIF_GUID = &H5     'Version6.0����
    NIIF_ICON_MASK = &HF    'Version6.0����
    NIIF_NOSOUND = &H10     'Version6.0��ֹ������Ӧ����
End Enum

Public Enum genumNotifyIconMouseEvent  '����¼�
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
    BalloonClick = (WM_USER + 5)
End Enum

Public gNotifyIconData As gtypeNOTIFYICONDATA


Public Enum genumFileTransimitType    '�ļ���������ö��
    ftSend = 1      '����
    ftReceive = 2   '����
End Enum

Public Enum genumSkinResChoose  '��������Դ�ļ�ѡ��
    sNone = 0   '��
    sMSVst = 1  'MicrosoftVista���
    sMSO7 = 2   'MicrosoftOffice2007���
    sMSO10 = 3   'MicrosoftOffice2010���
End Enum

Public Const gconAscAdd As Integer = 5      '�򵥼ӽ������ַ�ת��������
Public Const gconAddLenStart As Integer = 10    '�������Ŀ�ʼ���ֵ��ַ�����
Public Const gconSumLen As Integer = 60     '���ĵ����ַ���
Public Const gconMaxPWD As Integer = 20     '���������ַ���

'''�Զ��幫�ó���
Public Type gtypeCommonVariant
    TCPSetIP As String     '����IP��ַ
    TCPSetPort As Long     '���ö˿ں�
    TCPConnectMax As Long    '���������
    TCPDefaultIP As String      'Ĭ��IP��ַ
    TCPDefaultPort As Long      'Ĭ�϶˿ں�
    TCPWaitTime As Long     '����ȷ�ϵȴ�ʱ��
    
    TCPStateConnected As Boolean     '�ͻ������ӷ���˳ɹ���ʶ
    TCPStateServerStarted As Boolean '������������ʶ
    
    UpdatePCName As String  '�������³���ĵ�����
    UpdateAccount As String '�������³�����˺�
    UpdateUserName As String    '�������³�����û���
    
    FTChunkSize As Long   '�ļ�����ʱ�ķֿ��С
    FTWaitTime As Long    'ÿ���ļ�����ʱ�ĵȴ�ʱ�䣬��λ��
    
    EncryptKey As String    '���ܽ��ܵ���Կ
        
    ServerButtonStart As String       '������״̬����������
    ServerButtonClose As String       '�رշ���
    ServerStateError As String       '�쳣
    ServerStateStarted As String     '������
    ServerStateNotStarted As String  'δ����
    
    ClientStateConnected As String             '�ͻ���״̬��������
    ClientStateDisConnected As String          'δ����
    ClientStateConnectError As String          '�����쳣
    ClientButtonConnectToServer As String       '��������
    ClientButtonDisConnectFromServer As String  '�Ͽ�����
    
    PTFileName As String    'Э�飺�ļ�����ʶ
    PTFileSize As String    'Э�飺�ļ���С��ʶ
    PTFileFolder As String  'Э�飺�ļ�Ҫ������ļ�������ʶ
    PTFileStart As String   'Э�飺�ļ���ʼ�����ʶ
    PTFileEnd As String     'Э�飺�ļ����������ʶ
    PTFileSend As String    'Э�飺�ļ����ͱ�ʶ
    PTFileReceive As String 'Э�飺�ļ����ձ�ʶ
    PTFileExist As String  'Э�飺�ļ����ڱ�ʶ
    PTFileNoExist As String    'Э�飺�ļ������ڱ�ʶ
    
    PTVersionOfClient As String     'Э�飺�ͻ��˰汾��
    PTVersionNotUpdate As String    'Э�飺����Ҫ����
    PTVersionNeedUpdate As String   'Э�飺��Ҫ����
    
    PTClientConfirm As String   'Э�飺�ͻ���ȷ��
    PTClientIsTrue As String    'Э�飺�ͻ��˸�����˵�ȷ��
    
    PTDBDataSource As String    'Э�飺���ݿ��������ַ
    PTDBDatabase As String      'Э�飺���ݿ���
    PTDBUserID As String        'Э�飺���ݿ�����˺�
    PTDBPassword As String      'Э�飺���ݿ��������
    
    PTConnectIsFull As String   'Э�飺����������
    PTConnectTimeOut As String  'Э�飺��������ʱ�䵽
    
    PTClientUserComputerName As String  'Э�飺�ͻ��˼������
    PTClientUserLoginName As String 'Э�飺�ͻ����û���½��
    PTClientUserFullName As String  'Э�飺�ͻ����û�����
    
    EXENameOfClient As String   '�ͻ��˳���exe�ļ���
    EXENameOfUpdate As String   '���¶˳���exe�ļ���
    EXENameOfServer As String   '����˳���exe�ļ���
    EXENameOfSetup As String    '���°�װ��exe�ļ���
    
    CmdLineSeparator As String  '�����м����
    CmdLineParaOfHide As String '�����в���֮����
    
    '''SaveSetting(appname, section, key, setting)�����в���������
    '''GetSetting(appname, section, key[, default])
    
    RegAppName As String        'SaveSettin OR GetSetting������AppNameֵ
    RegSectionTCP As String     '����section_TCPֵ
    RegKeyTCPIP As String       '����key_IPֵ
    RegKeyTCPPort As String     '����key_portֵ
    
    RegSectionSkin As String    '����section_Skin
    RegKeySkinFile As String    '����Key_SkinFile
    
    RegSectionDBServer As String  '���ݿ��������Ϣ��
    RegKeyDBServerIP As String    '���ݿ������IP
    RegKeyDBServerDatabase As String    '���ݿ���
    RegKeyDBServerAccount As String   '���ݿ�����������˺�
    RegKeyDBServerPassword As String  '���ݿ��������������
        
    RegSectionUser As String    'Section_�û���Ϣ
    RegKeyUserLast As String    '����½�û���
    RegKeyUserList As String    '������½�����û����б�
    
    RegKeyCommandBars As String 'SaveCommandBars����RegistryKey
    RegKeyCBSServerSetting As String    'Server�ϵ�CBS�ؼ�ע����Ϣ����Keyֵ
    RegKeyCBSClientSetting As String    'Client�ϵ�CBS�ؼ�ע����Ϣ����Keyֵ
    RegKeyDockingPane As String
    RegKeyDockPaneServerSetting As String
    RegKeyDockPaneClientSetting As String
    
    RegSectionSettings As String    'Section_Settings��
    RegKeyServerWindowLeft As String  'Server��Key_����Leftֵ
    RegKeyServerWindowTop As String   '
    RegKeyServerWindowWidth As String '
    RegKeyServerWindowHeight As String    '
    RegKeyServerWindowStateMax As String
    RegKeyServerCommandbarsTheme As String    '
    
    
    RegKeyClientWindowLeft As String  'Client��Key_����Leftֵ
    RegKeyClientWindowTop As String   '
    RegKeyClientWindowWidth As String '
    RegKeyClientWindowHeight As String    '
    RegKeyClientWindowStateMax As String
    RegKeyClientCommandbarsTheme As String    '
    RegKeyClientTaskPanelTheme As String '�����˵�����
    RegKeyClientTaskPanelAutoFold As String '�����˵��Զ��۵�
    
    
    RegTrailPath As String  'ע�����HKEY_CURRENT_USER��SOFTWARE·��
    RegTrailKey As String   '������Ϣ-Keyֵ
    TrailPeriod As Long     '����������
    
    RegKeyParaWindowMinHide As String   '����Key-������С������
    RegKeyParaWindowCloseMin As String  '����Key-���ڵ���ر�ʱĬ����С��
    RegKeyParaAutoReStartServer As String   '������Ƿ��Զ���������
    RegKeyParaAutoStartupAtBoot As String   '�����Զ�����
    RegKeyParaLimitClientConnect As String  '���ƿͻ�������
    RegKeyParaLimitClientConnectTime As String '���ƿͻ�������ʱ��
    RegKeyParaLimitClientConnectNumber As String '���ƿͻ���������
    
    RegKeyParaRememberUserList As String  '��ס�û���
    RegKeyParaRememberUserPassword As String  '��ס����
    RegKeyParaUserAutoLogin As String '�Զ���½
    
    AppPath As String           'App·����ȷ������ַ�Ϊ"\"
    FolderNameTemp As String    '�ļ������ƣ�Temp��ȫ·��
    FolderNameData As String    '�ļ������ƣ�Data��ȫ·��
    FolderNameBin As String     '�ļ������ƣ�Bin��ȫ·��
    FolderBin As String     '�ļ������ƣ�Bin
    FolderData As String    '�ļ������ƣ�Data
    FolderTemp As String    '�ļ������ƣ�Temp
    
    FileNameErrLog As String    '�����¼��־�ļ���ȫ·��
    FileNameSkin As String      '������Դ�ļ���
    FileNameSkinIni As String   '���������ļ���
    FileNameLoginLog As String  '��½��־�ļ���
    
    UserAutoID As String    '�û���ʶID
    UserLoginName As String '�û���½��
    UserNickName As String  '�û��ǳ�
    UserFullName As String  '�û�����
    UserPassword As String  '�û�����
    UserDepartment As String    '�û����ڲ���
    UserLoginIP As String       '�û���½����IP
    UserComputerName As String  '�û���½��������
    
    rsURF As New ADODB.Recordset '�����û�������Ȩ����Ϣ
    
    AccountAdmin As String         '�ر��˺ţ�ϵͳ����Ա
    AccountSystem As String        '�ر��˺ţ�ϵͳ����Ա
    
    ConSource As String      '�������ݿ���������ƻ�IP��ַ
    ConUserID As String      '�������ݿ��û���
    ConPassword As String    '�������ݿ�����
    ConDatabase As String    '���ӵ����ݿ���
    ConString As String      '���ݿ������ַ���ȫ��
    
    FuncButton As String    '������𣺰�ť
    FuncForm As String      '������𣺴���
    FuncControl As String   '������������ؼ�
    FuncMainMenu As String  '����������˵�
    
    Formaty_M_dH_m_s As String  'ʱ���ʽyyyy-MM-dd HH:mm:ss
    Formatymdhms As String       'ʱ���ʽyyyyMMddHHmmss
    
    WindowWidth As Long     '����Ĭ�Ͽ��
    WindowHeight As Long    '����Ĭ�ϸ߶�
    
    CloseWindow As Boolean '�Ƿ������رմ��ڡ�����˵�Ƿ����˴������ϽǵĹرհ�ť
    ClientLoginShow As Boolean '��ʾ�ͻ��˵�½����
    ClientReLoad As Boolean   '���յ�����˷����������ͻ��˱�־
    ShowMainWindow As Boolean '�ͻ��˳ɹ���½����ʾ���������־
    UpdateRunOver As Boolean   '���³����Ƿ��������
    UnloadFromLogin As Boolean '�ӵ�½���ڴ������Ĺرճ���ָ��
    RestoreDBInfoOver As Boolean '���������ݿ�������Ϣ
    ClientLoginCheckOver As Boolean '��½�������
    ClientCancelAutoLogin As Boolean '��½�������ֶ���ʱȡ���Զ���½
    
    ParaBlnWindowStateMaxClient As Boolean '�����ϴιر�ʱ�Ƿ����
    ParaBlnWindowStateMaxServer As Boolean
    ParaBlnWindowMinHide As Boolean '��������С��ʱ�Ƿ�����
    ParaBlnWindowCloseMin As Boolean    '�����ڵ���رհ�ťʱ��С��
    ParaBlnAutoReStartServer As Boolean '����˳���Ͽ�����ʱ�Զ����¿�������
    ParaBlnAutoStartupAtBoot As Boolean '�����Զ�����
    ParaBlnLimitClientConnect As Boolean '���ƿͻ�������ʱ��
    ParaLimitClientConnectTime As Long  '���ƿͻ��������������ʱ���Ƕ���
    
    ParaBlnRememberUserList As Boolean  '��ס�û���
    ParaBlnRememberUserPassword As Boolean  '��ס����
    ParaBlnUserAutoLogin As Boolean '�Զ���½
    
End Type

Public Type gtypeFileTransmitVariant    '�Զ����ļ��������
    Connected As Boolean        'ȷ������״̬
    FileNumber As Integer       '�ļ�����ʱ�򿪵��ļ���
    FilePath As String          '�ļ�������ȫ·��
    FileName As String          '���ļ���������·��
    FileFolder As String        '�ļ��洢λ�õ��ļ������ƣ��ݲ�֧������·����Ĭ�϶���App.Path��
    FileSizeTotal As Long       '�ļ��ܴ�С
    FileSizeCompleted As Long   '�ļ��Ѵ����С
    FileTransmitState As Boolean    '�Ƿ��ڴ����ļ�
End Type

Public gVar As gtypeCommonVariant
Public gArr() As gtypeFileTransmitVariant

'''CommandBars��ID����
Public Type gtypeCommandBarID
    
    Sys As Long             'ģ��-ϵͳ
    
    SysLoginOut As Long     '�˳�ϵͳ
    SysLoginAgain As Long   '���µ�½
    SysAuthChangePassword As Long   '�����޸�
    SysAuthDepartment As Long       '���Ź���
    SysAuthUser As Long     '�û�����
    SysAuthLog As Long      '��־����
    SysAuthRole As Long     '��ɫ����
    SysAuthFunc As Long     '���ܹ���
    
    SysPrintMain As Long
    SysPrint As Long        '��ӡ
    SysPrintPageSet As Long '��ӡҳ������
    SysPrintPreview As Long '��ӡԤ��
    SysExportMain As Long
    SysExportToExcel As Long    '������Excel
    SysExportToWord As Long '������Word
    SysExportToText As Long '�������ı�
    SysExportToXML As Long  '����ΪXML�ĵ�
    SysExportToPDF As Long  '����ΪPDF
    SysExportToCSV As Long  '����ΪCSV�ļ�
    SysExportToHTML As Long '����ΪHTML
    
    SysSearch As Long   '���ڼ���������
    SysSearch1Label As Long
    SysSearch2TextBox As Long
    SysSearch3Button As Long
    SysSearch4ListBoxCaption As Long
    SysSearch4ListBoxFormID As Long
    SysSearch5Go As Long
    
    Help As Long        'ģ��-����
    
    HelpAbout As Long   '����
    HelpDocument As Long    '�����ĵ�
    HelpUpdate As Long  '������
    
    
    Wnd As Long 'ģ��-���ڿ���
    
    WndThemeSkinSet As Long '������������
    WndResetLayout As Long  '���ڲ�������
    
    WndOpenListCaption As Long '�Ѵ򿪴����б�
    WndOpenListID As Long
    
    WndToolBarCustomize As Long '�Զ��幤����
    WndToolBarList As Long '�������б�
    
    WndThemeCommandBars As Long '����������-CommandBars
    WndThemeCommandBarsOffice2000 As Long
    WndThemeCommandBarsOfficeXp As Long
    WndThemeCommandBarsOffice2003 As Long
    WndThemeCommandBarsWinXP As Long
    WndThemeCommandBarsWhidbey As Long
    WndThemeCommandBarsResource As Long
    WndThemeCommandBarsRibbon As Long
    WndThemeCommandBarsVS2008 As Long
    WndThemeCommandBarsVS6 As Long
    WndThemeCommandBarsVS2010 As Long
    
    WndThemeTaskPanel As Long   '�������(�����˵�)����-TaskPanel
    WndThemeTaskPanelOffice2000 As Long
    WndThemeTaskPanelOffice2003 As Long
    WndThemeTaskPanelNativeWinXP As Long
    WndThemeTaskPanelOffice2000Plain As Long
    WndThemeTaskPanelOfficeXPPlain As Long
    WndThemeTaskPanelOffice2003Plain As Long
    WndThemeTaskPanelNativeWinXPPlain As Long
    WndThemeTaskPanelToolbox As Long
    WndThemeTaskPanelToolboxWhidbey As Long
    WndThemeTaskPanelListView As Long
    WndThemeTaskPanelListViewOfficeXP As Long
    WndThemeTaskPanelListViewOffice2003 As Long
    WndThemeTaskPanelShortcutBarOffice2003 As Long
    WndThemeTaskPanelResource As Long
    WndThemeTaskPanelVisualStudio2010 As Long
    
    WndThemeSkin As Long    'ϵͳƤ������-SkinFrameWork
    WndThemeSkinCodejock As Long
    WndThemeSkinOffice2007 As Long
    WndThemeSkinOffice2010 As Long
    WndThemeSkinVista As Long
    WndThemeSkinWinXPRoyale As Long
    WndThemeSkinWinXPLuna As Long
    WndThemeSkinZune As Long
    
    WndSon As Long  '�Ӵ��ڿ���
    WndSonVbCascade As Long
    WndSonVbTileHorizontal As Long
    WndSonVbTileVertical As Long
    WndSonVbArrangeIcons As Long
    WndSonVbAllMin As Long
    WndSonVbAllBack As Long
    WndSonCloseAll As Long
    WndSonCloseCurrent As Long
    WndSonCloseLeft As Long
    WndSonCloseRight As Long
    WndSonCloseOther As Long
    
    
    Tool As Long    'ģ��--����
    
    toolOptions As Long  'ѡ��
    
    
'''**************************************************************'''

    
    Pane As Long   'ģ��--�������
    
    PaneNavi As Long '�����˵�
    
    PanePopupMenuNavi As Long   '�����˵�������嵯��ʽ�˵�ģ��
    PanePopupMenuNaviExpandALL As Long  'չ������
    PanePopupMenuNaviAutoFoldOther As Long  '�Զ��۵�����
    PanePopupMenuNaviFoldALL As Long    '�۵�����
    
    TabWorkspacePopupMenu As Long   '���ǩ�Ҽ��˵�ģ��
    
    StatusBarPane As Long               'ģ��-״̬�����
    
    StatusBarPaneProgress As Long       '״̬���н�����
    StatusBarPaneProgressText As Long   '״̬���н��Ȱٷ�ֵ
    StatusBarPaneUserInfo As Long       '״̬�����û���Ϣ
    StatusBarPaneTime As Long           '״̬����ʱ��
    StatusBarPaneConnectState As Long   '״̬��������״̬-Client
    StatusBarPaneConnectButton As Long  '״̬�������Ӱ�ť-Client
    StatusBarPaneServerState As Long    '״̬���з���������״̬-Server
    StatusBarPaneServerButton As Long   '״̬���з���������/�Ͽ�����ť-Server
    StatusBarPaneIP As Long     '״̬����IP
    StatusBarPanePort As Long   '״̬���ж˿�
    StatusBarPaneReStartButton As Long  '״̬���� �Զ���������ť
    
    IconPopupMenu As Long           '����ͼ��˵�
    IconPopupMenuShowWindow As Long '��ʾ����
    IconPopupMenuMinWindow As Long  '������С��
    IconPopupMenuMaxWindow As Long  '�������
    
End Type

Public Type gtypeValueAndErr    '���ڷ��ز���ֵ�Ĺ��̣�˳�㷵���쳣����
    Result As Boolean
    ErrNum As Long
End Type

Public Enum genumFileOpenType   '���ļ���ʽ
    udAppend    '��˳���ͷ��ʣ����ַ�׷�ӵ��ļ�
    udBinary    '�Զ����Ʒ���
    udInput     '��˳���ͷ��ʣ����ļ������ַ�
    udOutput    '��˳���ͷ��ʣ����ļ�����ַ�
    udRandom    '�������ʽ
End Enum

Public Enum genumFileWriteType  'д���ļ���ʽ
    udPut       '��Get����.For Binary��Random.
    udWrite     '��Input����
    udPrint     '��Line Input �� Input����
End Enum

Public Enum genumCharType   '�����ַ�����
    udUpperCase = 4     '����д��ĸ
    udLowerCase = 1     '��Сд��ĸ
    udNumber = 2        '������
    udUpperLowerNum = 7 '��д��Сд������
End Enum

Public Enum genumLogType    '������־��������ɾ���ġ���
    udSelect        '������ѯ
    udInsert
    udDelete
    udUpdate
    udSelectBatch   '�����ѯ
    udInsertBatch
    udDeleteBatch
    udUpdateBatch
End Enum

Public Enum genumGridExportType 'Flexcell Grid�����ļ�������
    fcCSV
    fcExcel
    fcHTML
    fcPDF
    fcXML
End Enum

Public Enum genumNumber '˳������
    eZero = 0
    eOne = 1
    eTwo = 2
    eThree = 3
    eFour = 4
    eFive = 5
End Enum


Public gID As gtypeCommandBarID '�������е�ȫ��CommandBars��ID����
Public gWind As Form            'ȫ������������




