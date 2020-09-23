;IFilter.asm  by Robert Rayment  18/12/01

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; VB
;' MCode Structure
;Public Type MCodeStruc
;   PICW As Long
;   PICH As Long
;   zmf1 As Single
;   zmf2 As Single
;   zdf As Single
;   zf As Single
;   ckQB As Long
;   ckABS As Long
;   QBLongColor As Long
;   PtrMat As Long
;   PtrPalBGR As Long
;End Type
;Public MC As MCodeStruc
;'-------------------------------------------------------
;
; res = CallWindowProc(ptMC, ptrStruc, 0&, 0&, 0&)
;
;' Determine PIC limits
;PicXLo = 2: PicXHi = PICW - 1
;PicYLo = 2: PicYHi = PICH - 1
;limx = 1
;limy = 1
;
;For ix = 1 To 5
;   If Mat(ix, 1) <> 0 Or Mat(ix, 5) <> 0 Then
;      PicXLo = 3: PicXHi = PICW - 2
;      limx = 2
;      Exit For
;   End If
;Next ix
;
;For iy = 1 To 5
;   If Mat(1, iy) <> 0 Or Mat(5, iy) <> 0 Then
;      PicYLo = 3: PicYHi = PICH - 2
;      limy = 2
;      Exit For
;   End If
;Next iy
;
;
;For iy = PicYLo To PicYHi 'Step 1 + 3 * Rnd
;For ix = PicXLo To PicXHi 'Step 1 + 3 * Rnd
;         
;   culB = 0
;   culG = 0
;   culR = 0
;   
;   For iyy = -limy To limy
;   miy = iyy + 3
;   For ixx = -limx To limx
;      
;      If ixx <> 0 Or iyy <> 0 Then
;         mix = ixx + 3
;         culB = culB + Mat(mix, miy) * PalBGR(1, ix + ixx, iy + iyy, 3)
;         culG = culG + Mat(mix, miy) * PalBGR(2, ix + ixx, iy + iyy, 3)
;         culR = culR + Mat(mix, miy) * PalBGR(3, ix + ixx, iy + iyy, 3)
;      End If
;      
;   Next ixx
;   Next iyy
;      
;   culB = (zmf1 * culB + zmf2 * Mat(3, 3) * PalBGR(1, ix, iy, 3)) / zdf
;   culG = (zmf1 * culG + zmf2 * Mat(3, 3) * PalBGR(2, ix, iy, 3)) / zdf
;   culR = (zmf1 * culR + zmf2 * Mat(3, 3) * PalBGR(3, ix, iy, 3)) / zdf
;      
;   If Form1.chkQB.Value = Checked Then
;      culB = culB + QBBlue
;      culG = culG + QBGreen
;      culR = culR + QBRed
;   Else
;      culB = culB + zf
;      culG = culG + zf
;      culR = culR + zf
;   End If
;      
;   If Form1.chkAbs.Value = Checked Then
;      culB = Abs(culB)
;      culG = Abs(culG)
;      culR = Abs(culR)
;   End If
;      
;   If culB > 255 Then culB = 255
;   If culG > 255 Then culG = 255
;   If culR > 255 Then culR = 255
;      
;   If culB < 0 Then culB = 0
;   If culG < 0 Then culG = 0
;   If culR < 0 Then culR = 0
;      
;   PalBGR(1, ix, iy, 2) = culB
;   PalBGR(2, ix, iy, 2) = culG
;   PalBGR(3, ix, iy, 2) = culR
;      
;Next ix
;Next iy
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

%macro movab 2		; name & num of parameters
  push dword %2		; 2nd param
  pop dword %1		; 1st param
%endmacro			; use  movab %1,%2
; Allows eg	movab bmW,[ebx+4]

%define PICW            [ebp-4]
%define PICH            [ebp-8]
%define zmf1            [ebp-12]
%define zmf2            [ebp-16]
%define zdf             [ebp-20]
%define zf              [ebp-24]
%define ckQB            [ebp-28]
%define ckABS           [ebp-32]
%define QBLongColor		[ebp-36]
%define PtrMat          [ebp-40]
%define PtrPalBGR		[ebp-44]

%define PicYLo		[ebp-48]
%define PicYHi		[ebp-52]
%define PicXLo		[ebp-56]
%define PicXHi		[ebp-60]
%define mlimy		[ebp-64]	; -limy
%define limy		[ebp-68]
%define mlimx		[ebp-72]	; -limx
%define limx		[ebp-76]
%define miy			[ebp-80]
%define mix			[ebp-84]


%define iy		[ebp-88]   ; PicYLo To PicYHi
%define ix		[ebp-92]   ; PicXLo To PicXHi
%define iyy		[ebp-96]   ; -limy ti limy 
%define ixx		[ebp-100]  ; -limx to limx

%define culB	[ebp-104]
%define culG	[ebp-108]
%define culR	[ebp-112]
%define culA	[ebp-116]

%define PalSize    [ebp-120]
%define LineBytes  [ebp-124]

%define QBculB	[ebp-128]
%define QBculG	[ebp-132]
%define QBculR	[ebp-136]
%define QBculA	[ebp-140]

%define MatVal	[ebp-144]
%define Temp	[ebp-148]
%define ptBGR2	[ebp-152]
%define ptBGR3	[ebp-156]

[bits 32]

	push ebp
	mov ebp,esp
	sub esp,156
	push edi
	push esi
	push ebx

	; Copy structure
	mov ebx,[ebp+8]
	
	movab PICW,          [ebx]
	movab PICH,          [ebx+4]
	movab zmf1,          [ebx+8]
	movab zmf2,          [ebx+12]
	movab zdf,           [ebx+16]
	movab zf,            [ebx+20]
	movab ckQB,          [ebx+24]
	movab ckABS,         [ebx+28]
	movab QBLongColor,   [ebx+32]
	movab PtrMat,        [ebx+36]
	movab PtrPalBGR,     [ebx+40]

	mov eax,PICH
	mov ebx,PICW
	mul ebx
	mov PalSize,eax		; In 4 byte chunks

	mov esi,PtrPalBGR   ; esi = pts to PalBGR(1,1,1,1)
	push esi
	pop edi
	mov eax,PalSize
	shl eax,2			; x4
	add edi,eax			; edi pts to PalBGR(1,1,1,2) Blue


;' Determine PIC limits
;PicXLo = 2: PicXHi = PICW - 1
;PicYLo = 2: PicYHi = PICH - 1
;limx = 1
;limy = 1
	
	mov eax,2
	mov PicXLo,eax
	mov PicYLo,eax
	dec eax
	mov limx,eax
	mov limy,eax
	mov eax,PICW
	dec eax
	mov PicXHi,eax
	mov eax,PICH
	dec eax
	mov PicYHi,eax
	
;For ix = 1 To 5
;   If Mat(ix, 1) <> 0 Or Mat(ix, 5) <> 0 Then
;      PicXLo = 3: PicXHi = PICW - 2
;      limx = 2
;      Exit For
;   End If
;Next ix

	mov ecx,5
nix:
	mov esi,PtrMat		; Mat(1,1)
	mov eax,ecx
	dec eax
	shl eax,2			; 4*(ix-1)
	add esi,eax			; Mat(ix,1)
	mov eax,[esi]
	cmp eax,0
	jne ReSizeXLims
	mov eax,80			; ie 20*4
	add esi,eax			; Mat(ix,5)
	mov eax,[esi]
	cmp eax,0
	jne ReSizeXLims
	dec ecx
	jnz nix
	jmp TestYLims
ReSizeXLims:
;      PicXLo = 3: PicXHi = PICW - 2
;      limx = 2
	
	mov eax,PicXLo
	inc eax
	mov PicXLo,eax
	mov eax,PicXHi
	dec eax
	mov PicXHi,eax
	mov eax,limx
	inc eax
	mov limx,eax
TestYLims:
;For iy = 1 To 5
;   If Mat(1, iy) <> 0 Or Mat(5, iy) <> 0 Then
;      PicYLo = 3: PicYHi = PICH - 2
;      limy = 2
;      Exit For
;   End If
;Next iy

	mov ecx,5
niy:
	mov esi,PtrMat		; Mat(1,1)
	mov eax,ecx
	dec eax
	mov ebx,20
	mul ebx
	add esi,eax			; Mat(1,iy)
	mov eax,[esi]
	cmp eax,0
	jne ReSizeYLims
	mov eax,16			; ie 4*4
	add esi,eax			; Mat(5,iy)
	mov eax,[esi]
	cmp eax,0
	jne ReSizeYLims
	dec ecx
	jnz niy
	jmp Setlimxy
ReSizeYLims:
;      PicYLo = 3: PicYHi = PICH - 2
;      limy = 2
	
	mov eax,PicYLo
	inc eax
	mov PicYLo,eax
	mov eax,PicYHi
	dec eax
	mov PicYHi,eax
	mov eax,limy
	inc eax
	mov limy,eax
Setlimxy:
	mov eax,limx
	neg eax
	mov mlimx,eax
	mov eax,limy
	neg eax
	mov mlimy,eax
	
	; Get QBRGB	
	mov eax,QBLongColor
	and eax,0FFh
	mov QBculR,eax
	mov eax,QBLongColor
	and eax,0FF00h
	shr eax,8
	mov QBculG,eax
	mov eax,QBLongColor
	and eax,0FF0000h
	shr eax,16
	mov QBculB,eax

;---------------------------------------------------  
	mov edi,PtrPalBGR
	mov eax,PalSize
	shl eax,2			; x4
	add edi,eax			; edi pts to PalBGR(1,1,1,2) Blue
	mov ptBGR2,edi
	add edi,eax			; edi pts to PalBGR(1,1,1,3) Blue
	mov ptBGR3,edi

;For iy = PicYLo To PicYHi 'Step 1 + 3 * Rnd
;For ix = PicXLo To PicXHi 'Step 1 + 3 * Rnd
;         
;   culB = 0
;   culG = 0
;   culR = 0
;   

	mov ecx,PicYLo
FORiy:
	mov iy,ecx
	push ecx
	
	mov ecx,PicXLo
FORix:
	mov ix,ecx
	push ecx
	;-----------------------------------------

	xor eax,eax
	mov culB,eax
	mov culG,eax
	mov culR,eax
	
	;   For iyy = -limy To limy
	;   miy = iyy + 3
	;   For ixx = -limx To limx
	
	mov ecx,mlimy
FORiyy:
	mov iyy,ecx
	push ecx
	
	mov eax,ecx
	add eax,3
	mov miy,eax
	
	mov ecx,mlimx
FORixx:
	mov ixx,ecx
	push ecx
	
	;      If ixx <> 0 Or iyy <> 0 Then
	;         mix = ixx + 3
	;         culB = culB + Mat(mix, miy) * PalBGR(1, ix + ixx, iy + iyy, 3)
	;         culG = culG + Mat(mix, miy) * PalBGR(2, ix + ixx, iy + iyy, 3)
	;         culR = culR + Mat(mix, miy) * PalBGR(3, ix + ixx, iy + iyy, 3)
	;      End If
	
	mov eax,ixx
	cmp eax,0
	jne AppFill
	mov eax,iyy
	cmp eax,0
	je near NEXixx
AppFill:
	mov eax,ixx
	add eax,3
	mov mix,eax
	
	mov edi,ptBGR3
	mov esi,PtrMat		; esi pts to Mat(1,1)
	
	; Get Mat(mix,miy) value
	mov eax,miy
	dec eax
	mov ebx,20
	mul ebx
	add esi,eax
	mov eax,mix
	dec eax
	shl eax,2
	add esi,eax
	mov eax,[esi]
	mov MatVal,eax
	
	;         culB = culB + Mat(mix, miy) * PalBGR(1, ix + ixx, iy + iyy, 3)
	;         culG = culG + Mat(mix, miy) * PalBGR(2, ix + ixx, iy + iyy, 3)
	;         culR = culR + Mat(mix, miy) * PalBGR(3, ix + ixx, iy + iyy, 3)
	
	mov edi,ptBGR3
	Call near GetAddrEDIixxxiyyy	; edi->PalBGR(1, ix + ixx, iy + iyy, 3)

	mov ebx,MatVal

	xor eax,eax
	mov AL,[edi]
	mul ebx
	add culB,eax
	xor eax,eax
	mov AL,[edi+1]
	mul ebx
	add culG,eax
	xor eax,eax
	mov AL,[edi+2]
	mul ebx
	add culR,eax

NEXixx:
	pop ecx
	inc ecx
	cmp ecx,limx
	jle near FORixx
NEXiyy:
	pop ecx
	inc ecx
	cmp ecx,limy
	jle near FORiyy
	
	;   culB = (zmf1 * culB + zmf2 * Mat(3, 3) * PalBGR(1, ix, iy, 3)) / zdf
	;   culG = (zmf1 * culG + zmf2 * Mat(3, 3) * PalBGR(2, ix, iy, 3)) / zdf
	;   culR = (zmf1 * culR + zmf2 * Mat(3, 3) * PalBGR(3, ix, iy, 3)) / zdf
	
	mov esi,PtrMat		; esi pts to Mat(1,1)

	; Get Mat(3,3) value
	mov eax,40
	add esi,eax
	mov eax,8
	add esi,eax
	mov eax,[esi]
	mov MatVal,eax
	
	mov edi,ptBGR3
	Call near GetAddrEDIixiy
	
	mov ebx,MatVal

	xor eax,eax
	mov AL,[edi]		; Blue
	mul ebx
	mov Temp,eax
	fild dword Temp
	fld dword zmf2
	fmulp st1			; zmf2 * Temp
	fild dword culB
	fld dword zmf1
	fmulp st1			; zmf1 * culB
	faddp st1
	fld dword zdf
	fdivp st1			; ( zmf1 * culB + zmf2 * Temp ) / zdf
	fistp dword culB
	
	xor eax,eax
	mov AL,[edi+1]		; Green
	mul ebx
	mov Temp,eax
	fild dword Temp
	fld dword zmf2
	fmulp st1			; zmf2 * Temp
	fild dword culG
	fld dword zmf1
	fmulp st1			; zmf1 * culG
	faddp st1
	fld dword zdf
	fdivp st1			; ( zmf1 * culG + zmf2 * Temp ) / zdf
	fistp dword culG
	
	xor eax,eax
	mov AL,[edi+2]		; Red
	mul ebx
	mov Temp,eax
	fild dword Temp
	fld dword zmf2
	fmulp st1			; zmf2 * Temp
	fild dword culR
	fld dword zmf1
	fmulp st1			; zmf1 * culR
	faddp st1
	fld dword zdf
	fdivp st1			; ( zmf1 * culR + zmf2 * Temp ) / zdf
	fistp dword culR



	;   If Form1.chkQB.Value = Checked Then
	;      culB = culB + QBBlue
	;      culG = culG + QBGreen
	;      culR = culR + QBRed
	;   Else
	;      culB = culB + zf
	;      culG = culG + zf
	;      culR = culR + zf
	;   End If

	mov eax,ckQB
	cmp eax,0
	je Addzf
	
	mov eax,culB
	add eax,QBculB
	mov culB,eax
	mov eax,culG
	add eax,QBculG
	mov culG,eax
	mov eax,culR
	add eax,QBculR
	mov culR,eax
	jmp CheckABS 

	;      culB = culB + zf
	;      culG = culG + zf
	;      culR = culR + zf

Addzf:
	fild dword culB
	fld dword zf
	faddp st1
	fistp dword culB
	
	fild dword culG
	fld dword zf
	faddp st1
	fistp dword culG

	fild dword culR
	fld dword zf
	faddp st1
	fistp dword culR

;   If Form1.chkAbs.Value = Checked Then
;      culB = Abs(culB)
;      culG = Abs(culG)
;      culR = Abs(culR)
;   End If

CheckABS:
	mov eax,ckABS
	cmp eax,0
	je CheckRange
	
	mov eax,culB
L1:	neg eax
	js L1
	mov culB,eax	; ABS(culB)
	
	mov eax,culG
L2:	neg eax
	js L2
	mov culG,eax	; ABS(culG)

	mov eax,culR
L3:	neg eax
	js L3
	mov culR,eax	; ABS(culR)


	;   If culB > 255 Then culB = 255
	;   If culG > 255 Then culG = 255
	;   If culR > 255 Then culR = 255
	;      
	;   If culB < 0 Then culB = 0
	;   If culG < 0 Then culG = 0
	;   If culR < 0 Then culR = 0

CheckRange:
	mov eax,culB
	cmp eax,255
	jl B255
	mov eax,255
	jmp STOculB
B255:
	cmp eax,0
	jge STOculB
	xor eax,eax
STOculB:
	mov culB,eax

	mov eax,culG
	cmp eax,255
	jl G255
	mov eax,255
	jmp STOculG
G255:
	cmp eax,0
	jge STOculG
	xor eax,eax
STOculG:
	mov culG,eax

	mov eax,culR
	cmp eax,255
	jl R255
	mov eax,255
	jmp STOculR
R255:
	cmp eax,0
	jge STOculR
	xor eax,eax
STOculR:
	mov culR,eax

;   PalBGR(1, ix, iy, 2) = culB
;   PalBGR(2, ix, iy, 2) = culG
;   PalBGR(3, ix, iy, 2) = culR

	;-----------------------------------------

	mov edi,ptBGR2
	Call near GetAddrEDIixiy
	
	mov eax,culB
	mov [edi],AL
	mov eax,culG
	mov [edi+1],AL
	mov eax,culR
	mov [edi+2],AL

NEXix:
	pop ecx
	inc ecx
	cmp ecx,PicXHi
	jle near FORix
NEXiy:
	pop ecx
	inc ecx
	cmp ecx,PicYHi
	jle near FORiy

GETOUT:
	pop ebx
	pop esi
	pop edi
	mov esp,ebp
	pop ebp
	ret 16

;============================================================

GetAddrEDIixxxiyyy: ; In edi, ix + ixx, iy + iyy  Out: new edi->B
                ; Uses eax,ebx
	;B = edi + 4 * [(iy + iyy - 1) * PICW + (ix + ixx - 1))]
	mov eax,iy
	mov ebx,iyy
	add eax,ebx
	dec eax
	mov ebx,PICW
	mul ebx
	push eax
	
	mov eax,ix
	mov ebx,ixx
	add eax,ebx
	dec eax

	pop ebx
	add eax,ebx
	shl eax,2		; x4
	add edi,eax
RET
;============================================================
GetAddrEDIixiy: ; In edi, ix, iy  Out: new edi->B
                ; Uses eax,ebx
	;B = edi + 4 * [(iy - 1) * PICW + (ix - 1))]
	mov eax,iy
	dec eax
	mov ebx,PICW
	mul ebx
	push eax
	
	mov eax,ix
	dec eax

	pop ebx
	add eax,ebx
	shl eax,2		; x4
	add edi,eax
RET
;============================================================

;############################################################
