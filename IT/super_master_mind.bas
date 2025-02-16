REM  *****  BASIC  *****

Sub Main
With ThisComponent.Sheets(0)
	for Y1= 5 to 27 step 2
		If .getCellByPosition(4,Y1).String=ThisComponent.Sheets(1).getCellRangeByName("E3").String and .getCellByPosition(5,Y1).String=ThisComponent.Sheets(1).getCellRangeByName("F3").String and .getCellByPosition(6,Y1).String=ThisComponent.Sheets(1).getCellRangeByName("G3").String and .getCellByPosition(7,Y1).String=ThisComponent.Sheets(1).getCellRangeByName("H3").String and .getCellByPosition(8,Y1).String=ThisComponent.Sheets(1).getCellRangeByName("I3").String then
			msgbox("Hai vinto," & CHR$(13) & "Vuoi iniziare un'altra partita?" & CHR$(13) & "Clicca su NUOVA PARTITA.",,"Vittoria")
			.getCellRangeByName("E3").String=ThisComponent.Sheets(1).getCellRangeByName("E3").String
			.getCellRangeByName("F3").String=ThisComponent.Sheets(1).getCellRangeByName("F3").String
			.getCellRangeByName("G3").String=ThisComponent.Sheets(1).getCellRangeByName("G3").String
			.getCellRangeByName("H3").String=ThisComponent.Sheets(1).getCellRangeByName("H3").String
			.getCellRangeByName("I3").String=ThisComponent.Sheets(1).getCellRangeByName("I3").String
			exit for
		end if
	next Y1 
 
	for y= 5 to 27 step 2 
	
		if .getCellByPosition(4,y).String<>"" and .getCellByPosition(5,y).String<>"" and .getCellByPosition(6,y).String<>"" and .getCellByPosition(7,y).String<>"" and .getCellByPosition(8,y).String<>"" then
  
			for x= 4 to 8 'input'
				'neri'
				if .getCellByPosition(x,y).String=ThisComponent.Sheets(1).getCellByPosition(x,2).String  then
					.getCellByPosition(1,y).value=.getCellByPosition(1,y).value+1 'contatore'
					ThisComponent.Sheets(1).getCellByPosition(x,y).value=1 'annotazione dell'algoritmo'
					goto avanti 'il nero non deve passare nel bianco'
				end if 
			for xi= 4 to 8 'output'
				'bianchi'
				if .getCellByPosition(x,y).String=ThisComponent.Sheets(1).getCellByPosition(xi,2).String and ThisComponent.Sheets(1).getCellByPosition(xi,y).value<>1 then
					.getCellByPosition(3,y).value=.getCellByPosition(3,y).value+1 'contatore'
					ThisComponent.Sheets(1).getCellByPosition(xi,y).value=1 'annotazione dell'algoritmo'
					exit for 'uscita, ogni input ha un solo output'
				end if 
			next xi
			avanti: 'uscita dal ciclo più interno'
 			next x
  
			exit for 'esci per incrementare le prestazioni'
		end if
 
	next y
 
end with
End Sub

Rem *************************************************************************************************************************************************************************************************************************

Sub Riavvia
With ThisComponent.Sheets(1)
 
'ciclo per nascondere la barra di soluzione nella pagina di gioco' 
	for counter1= 4 to 8
		ThisComponent.Sheets(0).getCellByPosition(counter1,2).string="?"
	next counter1

'azzera i responsi (contatori)'
	for counter2= 5 to 27 step 2
		ThisComponent.Sheets(0).getCellByPosition(1,counter2).string=""
		ThisComponent.Sheets(0).getCellByPosition(3,counter2).string=""
	next counter2

'azzera i codici del giocatore'
	for counter3= 5 to 27 step 2
 		for counter4= 4 to 8
			ThisComponent.Sheets(0).getCellByPosition(counter4,counter3).string=""
		next counter4
	next counter3

'azzera le annotazioni dell'algoritmo'
	for counter5= 5 to 27 step 2
		for counter6= 4 to 8
			.getCellByPosition(counter6,counter5).string=""
		next counter6
	next counter5

'mescolatore casuale dei colori nella soluzione (bot)'
	dim a
	a=int( Rnd()(0) * (7))
	Dim A1 As Variant
	A1= Array("Giallo","Rosso","Arancione","Verde","Blu","Beige","Bianco","Nero")
	dim rnd_a
	rnd_a= A1(a)
 
	dim b
	b=int( Rnd()(0) * (7))
	Dim B1 As Variant
	B1=Array("Giallo","Rosso","Arancione","Verde","Blu","Beige","Bianco","Nero")
	dim rnd_b
	rnd_b= B1(b)
 
	dim c
	c=int( Rnd()(0) * (7))
	Dim C1 As Variant
	C1=Array("Giallo","Rosso","Arancione","Verde","Blu","Beige","Bianco","Nero")
	dim rnd_c
	rnd_c= C1(c)
	  
	dim d
	d=int( Rnd()(0) * (7))
	Dim D1 As Variant
	D1=Array("Giallo","Rosso","Arancione","Verde","Blu","Beige","Bianco","Nero")
	dim rnd_d
	rnd_d= D1(d)
	 
	dim e
	e=int( Rnd()(0) * (7))
	Dim E1 As Variant
	E1=Array("Giallo","Rosso","Arancione","Verde","Blu","Beige","Bianco","Nero")
	dim rnd_e
	rnd_e= E1(e)

'main mescolatore'
	.getCellRangeByName("E3").string=rnd_a
	.getCellRangeByName("F3").string=rnd_b
	.getCellRangeByName("G3").string=rnd_c
	.getCellRangeByName("H3").string=rnd_d
	.getCellRangeByName("I3").string=rnd_e

end with
End Sub
