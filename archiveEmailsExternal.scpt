FasdUAS 1.101.10   ��   ��    k             l    � ����  O     �  	  k    � 
 
     l   ��  ��    $  List of folder IDs to process     �   <   L i s t   o f   f o l d e r   I D s   t o   p r o c e s s      r        J    
       m    ���� �      m    ���� �      m    ���� �   ��  m    ���� ���    o      ���� 0 	folderids 	folderIDs      l   ��������  ��  ��        l   ��   !��     / ) Iterate through each specified folder ID    ! � " " R   I t e r a t e   t h r o u g h   e a c h   s p e c i f i e d   f o l d e r   I D   #�� # X    � $�� % $ k    � & &  ' ( ' r    % ) * ) e    # + + 5    #�� ,��
�� 
cMFo , o     ���� 0 folderid folderID
�� kfrmID   * o      ���� 0 currentfolder currentFolder (  - . - l  & &��������  ��  ��   .  / 0 / l  & &�� 1 2��   1 - ' Check if there are messages to forward    2 � 3 3 N   C h e c k   i f   t h e r e   a r e   m e s s a g e s   t o   f o r w a r d 0  4�� 4 Z   & � 5 6�� 7 5 ?   & / 8 9 8 l  & - :���� : I  & -�� ;��
�� .corecnte****       **** ; n  & ) < = < 2  ' )��
�� 
msg  = o   & '���� 0 currentfolder currentFolder��  ��  ��   9 m   - .����   6 k   2 � > >  ? @ ? r   2 : A B A n   2 8 C D C 4  5 8�� E
�� 
cobj E m   6 7����  D n  2 5 F G F 2  3 5��
�� 
msg  G o   2 3���� 0 currentfolder currentFolder B o      ���� 0 firstmessage firstMessage @  H I H l  ; ;��������  ��  ��   I  J K J l  ; ;�� L M��   L   Create a forward message    M � N N 2   C r e a t e   a   f o r w a r d   m e s s a g e K  O P O r   ; F Q R Q I  ; B�� S T
�� .OEMamFwdnull���     cEvt S o   ; <���� 0 firstmessage firstMessage T �� U��
�� 
ropw U m   = >��
�� boovfals��   R o      ���� 0 
newmessage 
newMessage P  V�� V O   G � W X W k   M � Y Y  Z [ Z l  M M�� \ ]��   \ &   Set the subject with FWD prefix    ] � ^ ^ @   S e t   t h e   s u b j e c t   w i t h   F W D   p r e f i x [  _ ` _ r   M \ a b a b   M V c d c m   M P e e � f f  A R C H I V E :   d n  P U g h g 1   Q U��
�� 
subj h o   P Q���� 0 firstmessage firstMessage b 1   V [��
�� 
subj `  i j i l  ] ]��������  ��  ��   j  k l k l  ] ]�� m n��   m / ) Set the recipient to the specified email    n � o o R   S e t   t h e   r e c i p i e n t   t o   t h e   s p e c i f i e d   e m a i l l  p q p I  ] ~���� r
�� .corecrel****      � null��   r �� s t
�� 
kocl s m   _ b��
�� 
rcpt t �� u v
�� 
insh u o   e h���� 0 
newmessage 
newMessage v �� w��
�� 
prdt w K   k x x x �� y��
�� 
emad y K   n v z z �� {��
�� 
radd { m   q t | | � } } " y o u n a l y o n s @ m e . c o m��  ��  ��   q  ~  ~ l   ��������  ��  ��     � � � l   �� � ���   �   Send the message    � � � � "   S e n d   t h e   m e s s a g e �  ��� � I   �������
�� .mailsendnull���     msg ��  ��  ��   X o   G J���� 0 
newmessage 
newMessage��  ��   7 k   � � � �  � � � l  � ��� � ���   � < 6 Notify if no messages are found in the current folder    � � � � l   N o t i f y   i f   n o   m e s s a g e s   a r e   f o u n d   i n   t h e   c u r r e n t   f o l d e r �  ��� � I  � ��� ���
�� .sysodlogaskr        TEXT � b   � � � � � m   � � � � � � � J N o   m e s s a g e s   t o   f o r w a r d   i n   f o l d e r   I D :   � o   � ����� 0 folderid folderID��  ��  ��  �� 0 folderid folderID % o    ���� 0 	folderids 	folderIDs��   	 m      � ��                                                                                  OPIM  alis    N  Macintosh HD                   BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  "Applications/Microsoft Outlook.app  / ��  ��  ��     ��� � l     ��������  ��  ��  ��       �� � ���   � ��
�� .aevtoappnull  �   � **** � �� ����� � ���
�� .aevtoappnull  �   � **** � k     � � �  ����  ��  ��   � ���� 0 folderid folderID �  ����������������������������������� e������������ |������ ����� ��� ��� ��� ��� �� 0 	folderids 	folderIDs
�� 
kocl
�� 
cobj
�� .corecnte****       ****
�� 
cMFo
�� kfrmID  �� 0 currentfolder currentFolder
�� 
msg �� 0 firstmessage firstMessage
�� 
ropw
�� .OEMamFwdnull���     cEvt�� 0 
newmessage 
newMessage
�� 
subj
�� 
rcpt
�� 
insh
�� 
prdt
�� 
emad
�� 
radd�� 
�� .corecrel****      � null
�� .mailsendnull���     msg 
�� .sysodlogaskr        TEXT�� �� ������vE�O ��[��l 	kh  *��0EE�O��-j 	j X��-�k/E�O��fl E` O_  9a �a ,%*a ,FO*�a a _ a a a a lla  O*j UY a �%j [OY��Uascr  ��ޭ