Attribute VB_Name = "MODChat2Pcks"
Option Explicit

Public Function Authentication(UserName As String, PassWord As String) As String
Dim User As String
  User = Replace(UserName, " ", "%20")
  Authentication = "GET /config/login?.tries=1&.src=chat&.last=&promo=&lg=us&.intl=us&.bypass=&.chkP=Y&.done=http%3A%2F%2Fchat.yahoo.com%2F%3Froom%3D30%2527s%253a%253a1600326617&login=" & User & "&passwd=" & PassWord & "&n=1 HTTP/1.0" & _
    vbCrLf & "Accept: */*" & vbCrLf & "Accept: text/html" & vbCrLf & vbCrLf
End Function

Public Function LoginChat2(UserName As String, Cookie As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "0À€" & UserName & "À€1À€" & UserName & "À€6À€" & Cookie & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  LoginChat2 = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H1E) & Chr(&H5A) & Chr(&H55) & Chr(&HAA) & Chr(&H55) & String(4, 0) & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 44                     YMSG.....D
'00 A8 00 00 00 00 EE 01-C7 EF 31 C0 80 6A 61 79   .¨....î.Çï1À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 31 30 34-C0 80 47 65 6E 65 72 61   _upÀ€104À€Genera
'74 69 6F 6E 20 58 3A 32-34 C0 80 31 31 37 C0 80   tion X:24À€117À€
'62 6C 61 68 C0 80 31 32-34 C0 80 31 C0 80         blahÀ€124À€1À€
Public Function Chat2ChatSend(UserName As String, RoomName As String, Message As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€104À€" & RoomName & "À€117À€" & Message & "À€124À€1À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  Chat2ChatSend = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&HA8) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 35                     YMSG.....5
'00 20 5A 55 AA 55 EB 19-F2 F6 30 C0 80 64 72 61   . ZUªUë.òö0À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 C0 80 64 72   ma_stinksÀ€1À€dr
'61 6D 61 5F 73 74 69 6E-6B 73 C0 80 35 C0 80 71   ama_stinksÀ€5À€q
'71 2E 37 38 39 C0 80 31-34 C0 80 79 6F C0 80      q.789À€14À€yoÀ€
Public Function Chat2PmSend(UserName As String, WhoTo As String, Message As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "0À€" & UserName & "À€1À€" & UserName & "À€5À€" & WhoTo & "À€14À€" & Message & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  Chat2PmSend = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H20) & Chr(&H5A) & Chr(&H55) & Chr(&HAA) & Chr(&H55) & Session & Pck
End Function

''''''
Public Function LoginWebMessenger(UserName As String, Cookie As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "0À€" & UserName & "À€1À€" & UserName & "À€6À€" & Cookie & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  LoginWebMessenger = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H2) & Chr(&H26) & Chr(&H5A) & Chr(&H55) & Chr(&HAA) & Chr(&H55) & String(4, 0) & Pck
End Function

'59 4D-53 47 00 65 00 00 00 4B                     YMSG.e...K
'00 06 5A 55 AA 55 F2 C5-29 9F 30 C0 80 6A 61 79   ..ZUªUòÅ)Ÿ0À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 31 C0 80-6A 61 79 5F 64 6F 67 5F   _upÀ€1À€jay_dog_
'74 68 61 74 73 5F 77 68-61 74 73 5F 75 70 C0 80   thats_whats_upÀ€
'35 C0 80 71 71 2E 37 38-39 C0 80 31 34 C0 80 79   5À€qq.789À€14À€y
'6F 6F 6F C0 80                                    oooÀ€
Public Function WebMsgPmSend(UserName As String, WhoTo As String, Message As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "0À€" & UserName & "À€1À€" & UserName & "À€5À€" & WhoTo & "À€14À€" & Message & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgPmSend = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H6) & Chr(&H5A) & Chr(&H55) & Chr(&HAA) & Chr(&H55) & Session & Pck
End Function

'59 4D-53 47 00 65 00 00 00 42                     YMSG.e...B
'00 86 00 00 00 00 F2 C4-3F 13 31 C0 80 6A 61 79   .†....òÄ?.1À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 37 C0 80-71 71 2E 37 38 39 C0 80   _upÀ€7À€qq.789À€
'31 34 C0 80 54 68 61 6E-6B 73 2C 20 62 75 74 20   14À€Thanks, but
'6E 6F 20 74 68 61 6E 6B-73 2E C0 80               no thanks.À€
Public Function WebMsgDenyAdd(UserName As String, WhoTo As String, Message As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€7À€" & WhoTo & "À€14À€" & Message & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgDenyAdd = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H86) & Chr(&H5A) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 65 00 00 00 96                     YMSG.e...–
'00 84 00 00 00 00 F2 C4-3F 13 31 C0 80 6A 61 79   .„....òÄ?.1À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 37 C0 80-6A 61 79 73 5F 62 6F 74   _upÀ€7À€jays_bot
'34 C0 80 36 35 C0 80 46-72 69 65 6E 64 73 C0 80   4À€65À€FriendsÀ€
'32 32 37 C0 80 75 35 64-32 6E 42 C0 80 32 32 36   227À€u5d2nBÀ€226
'C0 80 75 76 5F 42 6E 75-56 5A 46 65 6D 42 42 52   À€uv_BnuVZFemBBR
'61 6C 6D 76 4C 5F 78 6E-35 31 53 4C 65 71 67 66   almvL_xn51SLeqgf
'4C 77 39 55 35 42 6B 67-44 61 46 2E 72 4C 50 55   Lw9U5BkgDaF.rLPU
'67 6B 47 65 77 75 56 45-4C 39 4C 6C 79 6E 59 71   gkGewuVEL9LlynYq
'68 4A 59 6B 35 65 71 77-5A 62 6E 47 49 2D C0 80   hJYk5eqwZbnGI-À€
Public Function WebMsgRemoveUser(UserName As String, WhoTo As String, Group As String, AuthWord As String, Token As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€7À€" & WhoTo & "À€65À€" & Group & "À€227À€" & AuthWord & "À€226À€" & Token & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgRemoveUser = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H84) & String(4, 0) & Session & Pck
End Function

'invisible
'59 4D-53 47 00 65 00 00 00 08    YMSG.e....
'00 03 00 00 00 00 F2 D1-FC EF 31 30 C0 80 31 32   ......òÑüï10À€12
'C0 80                                             À€
Public Function WebMsgInvisible(Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "10À€12À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgInvisible = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H3) & String(4, 0) & Session & Pck
End Function

'visible
'59 4D-53 47 00 65 00 00 00 02    YMSG.e....
'00 04 00 00 00 00 F2 D1-FC EF C0 80               ......òÑüïÀ€
Public Function WebMsgVisible(Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgVisible = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H4) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 65 00 00 00 02    YMSG.e....
'02 28 00 00 00 00 F2 C0-03 9A C0 80               .(....òÀ.šÀ€
Public Function WebMsgAddPrompt(Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgAddPrompt = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H2) & Chr(&H28) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 65 00 00 00 84                     YMSG.e...„
'00 83 00 00 00 00 F2 C0-03 9A 31 C0 80 64 72 61   .ƒ....òÀ.š1À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 37 C0 80 71 71   ma_stinksÀ€7À€qq
'2E 37 38 39 C0 80 31 34-C0 80 61 64 64 20 6D 65   .789À€14À€add me
'C0 80 36 35 C0 80 46 72-69 65 6E 64 73 C0 80 32   À€65À€FriendsÀ€2
'32 37 C0 80 65 65 32 54-46 C0 80 32 32 36 C0 80   27À€ee2TFÀ€226À€
'6F 69 72 56 41 65 56 5A-46 65 6D 32 4A 43 65 30   oirVAeVZFem2JCe0
'61 5A 78 35 43 41 4F 4F-46 33 71 41 34 45 69 56   aZx5CAOOF3qA4EiV
'34 61 73 2E 43 6D 64 62-54 79 6C 5A 43 6B 33 43   4as.CmdbTylZCk3C
'52 4B 63 37 69 44 77 35-57 45 2E 5F C0 80         RKc7iDw5WE._À€
Public Function WebMsgAddUser(UserName As String, WhoTo As String, Message As String, List As String, AuthWord As String, Token As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€7À€" & WhoTo & "À€14À€" & Message & "À€65À€" & List & "À€227À€" & AuthWord & "À€226À€" & Token & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  WebMsgAddUser = "YMSG" & YBuild(2) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H83) & String(4, 0) & Session & Pck
End Function
''''

'59 4D-53 47 00 00 00 00 00 23                     YMSG.....#
'00 1E 00 00 00 00 EA 03-B0 CD 30 C0 80 64 72 61   ......ê.°Í0À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 C0 80 64 72   ma_stinksÀ€1À€dr
'61 6D 61 5F 73 74 69 6E-6B 73 C0 80 38            ama_stinksÀ€8

'<
'59 4D-53 47 00 00 00 00 00 AD                     YMSG.....­
'00 55 00 00 00 00 EA 03-B0 CD 38 37 C0 80 0A C0   .U....ê.°Í87À€.À
'80 38 38 C0 80 C0 80 38-39 C0 80 64 72 61 6D 61   €88À€À€89À€drama
'5F 73 74 69 6E 6B 73 2C-64 72 61 6D 61 5F 61 73   _stinks,drama_as
'73 2C 6A 65 73 75 73 5F-68 61 74 65 73 5F 64 72   s,jesus_hates_dr
'61 6D 61 2C 63 79 62 65-72 5F 64 72 61 6D 61 2C   ama,cyber_drama,
'64 72 61 6D 61 5F 6E 61-6D 65 2C 6E 6F 5F 6F 6E   drama_name,no_on
'65 5F 6C 69 6B 65 73 5F-61 5F 64 72 61 6D 61 5F   e_likes_a_drama_
'71 75 65 65 6E 2C 64 72-61 6D 61 5F 65 62 6F 6E   queen,drama_ebon
'69 63 73 C0 80 31 35 33-C0 80 31 C0 80 39 30 C0   icsÀ€153À€1À€90À
'80 30 C0 80 33 C0 80 C0-80 39 33 C0 80 30 C0 80   €0À€3À€À€93À€0À€
'31 38 36 C0 80 1A A8 02-C0 80 32 31 37 C0 80 31   186À€.¨.À€217À€1
'37 33 30 35 30 C0 80                              73050À€

'59 4D-53 47 00 0A 00 00 00 30                     YMSG.....0
'00 96 00 00 00 00 EA 03-B0 CD 31 C0 80 64 72 61   .–....ê.°Í1À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 36 C0 80 61 62   ma_stinksÀ€6À€ab
'63 64 65 66 C0 80 39 38-C0 80 75 73 C0 80 31 33   cdefÀ€98À€usÀ€13
'35 C0 80 64 63 31 32 35-C0 80                     5À€dc125À€
Public Function LoginToChat(UserName As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€6À€abcdefÀ€98À€usÀ€135À€dc125À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  LoginToChat = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H96) & String(4, 0) & Session & Pck
End Function
'<
'59 4D-53 47 00 00 00 00 00 23                              YMSG.....#
'00 96 00 00 00 00 EA 03-B0 CD 30 C0 80 64 72 61   .–....ê.°Í0À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 C0 80 64 72   ma_stinksÀ€1À€dr
'61 6D 61 5F 73 74 69 6E-6B 73 C0 80 38            ama_stinksÀ€8

'59 4D-53 47 00 0A 00 00 00 32                     YMSG.....2
'00 98 00 00 00 00 EA 03-B0 CD 31 C0 80 64 72 61   .˜....ê.°Í1À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 36 32 C0 80 32   ma_stinksÀ€62À€2
'C0 80 31 30 34 C0 80 47-65 6E 65 72 61 74 69 6F   À€104À€Generatio
'6E 20 58 C0 80 31 32 39-C0 80 C0 80               n XÀ€129À€À€
Public Function JoinRoom(UserName As String, RoomName As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€62À€2À€104À€" & RoomName & "À€129À€À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  JoinRoom = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H98) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 22                     YMSG....."
'00 1F 00 00 00 00 EA 0D-2C AA 30 C0 80 64 72 61   ......ê.,ª0À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 C0 80 64 72   ma_stinksÀ€1À€dr
'61 6D 61 5F 73 74 69 6E-6B 73 C0 80               ama_stinksÀ€

'59 4D-53 47 00 0A 00 00 00 2D                     YMSG.....-
'00 9B 00 00 00 00 EB 1F-FD 24 31 C0 80 6A 61 79   .›....ë.ý$1À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 31 30 34-C0 80 51 51 27 73 20 52   _upÀ€104À€QQ's R
'6F 6F 6D 3A 31 C0 80                              oom:1À€
Public Function LogOutChat2(UserName As String, RoomName As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€104À€" & RoomName & "À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  LogOutChat2 = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H9B) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 40                     YMSG.....@
'00 A8 00 00 00 00 EB 1F-FD 24 31 C0 80 6A 61 79   .¨....ë.ý$1À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 31 30 34-C0 80 51 51 27 73 20 52   _upÀ€104À€QQ's R
'6F 6F 6D 3A 31 C0 80 31-31 37 C0 80 20 79 6F 6F   oom:1À€117À€ yoo
'C0 80 31 32 34 C0 80 32-C0 80                     À€124À€2À€
Public Function Chat2EmoteChatSend(UserName As String, RoomName As String, Message As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€104À€" & RoomName & "À€117À€" & Message & "À€124À€2À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  Chat2EmoteChatSend = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&HA8) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 48                     YMSG.....H
'00 A8 00 00 00 00 EB 1F-FD 24 31 C0 80 6A 61 79   .¨....ë.ý$1À€jay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 31 30 34-C0 80 51 51 27 73 20 52   _upÀ€104À€QQ's R
'6F 6F 6D 3A 31 C0 80 31-31 37 C0 80 2E 20 6F 20   oom:1À€117À€. o
'4F 20 28 20 79 6F 20 29-C0 80 31 32 34 C0 80 32   O ( yo )À€124À€2
'C0 80                                             À€
Public Function Chat2ThinkChatSend(UserName As String, RoomName As String, Message As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€104À€" & RoomName & "À€117À€. o O ( " & Message & " )À€124À€2À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  Chat2ThinkChatSend = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&HA8) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 35                     YMSG.....5
'00 A8 00 00 00 00 E8 1F-F6 4E 31 C0 80 64 72 61   .¨....è.öN1À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 30 34 C0 80   ma_stinksÀ€104À€
'65 6E 65 77 72 3A 31 C0-80 31 31 37 C0 80 69 73   enewr:1À€117À€is
'20 62 61 63 6B C0 80 31-32 34 C0 80 32 C0 80       backÀ€124À€2À€
Public Function Available(UserName As String, RoomName As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€104À€" & RoomName & "À€117À€is backÀ€124À€2À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  Available = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&HA8) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 3C                     YMSG.....<
'00 A8 00 00 00 00 E8 1F-F6 4E 31 C0 80 64 72 61   .¨....è.öN1À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 30 34 C0 80   ma_stinksÀ€104À€
'65 6E 65 77 72 3A 31 C0-80 31 31 37 C0 80 69 73   enewr:1À€117À€is
'20 61 77 61 79 20 28 42-75 73 79 29 C0 80 31 32    away (Busy)À€12
'34 C0 80 32 C0 80                                 4À€2À€
Public Function Busy(UserName As String, RoomName As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€104À€" & RoomName & "À€117À€is away (Busy)À€124À€2À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  Busy = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&HA8) & String(4, 0) & Session & Pck
End Function

'59 4D-53 47 00 0A 00 00 00 42                     YMSG.....B
'00 83 00 00 00 00 E8 13-5E 39 31 C0 80 64 72 61   .ƒ....è.^91À€dra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 37 C0 80 71 71   ma_stinksÀ€7À€qq
'2E 37 38 39 C0 80 31 34-C0 80 C0 80 36 35 C0 80   .789À€14À€À€65À€
'43 68 61 74 20 46 72 69-65 6E 64 73 C0 80 32 32   Chat FriendsÀ€22
'37 C0 80 C0 80 32 32 36-C0 80 C0 80               7À€À€226À€À€
Public Function AddUser(UserName As String, WhoTo As String, Message As String, List As String, Session As String) As String
Dim x(1 To 2) As Integer
Dim Pck As String
  Pck = "1À€" & UserName & "À€7À€" & WhoTo & "À€14À€" & Message & "À€65À€" & List & "À€227À€À€226À€À€"
  x(1) = 0
  x(2) = Len(Pck)
ReCheck:
    If x(2) > 255 Then
      x(2) = x(2) - 256
      x(1) = x(1) + 1
      GoTo ReCheck
    End If
  AddUser = "YMSG" & YBuild(1) & String(2, 0) & Chr(x(1)) & Chr(x(2)) & Chr(&H0) & Chr(&H83) & String(4, 0) & Session & Pck
End Function
