system "cls"

cell_b1 = "Załącznik nr 2c - Wzór formularza oceny - ocena merytoryczna II stopnia"
cell_b2 = "FORMULARZ OCENY - OCENA MERYTORYCZNA
		Centrum Projektów Polska Cyfrowa
		w ramach oceny II stopnia
		I Oś priorytetowa – Powszechny dostęp do szybkiego internetu
		Działanie 1.1 – Wyeliminowanie terytorialnych różnic w możliwości dostępu do szerokopasmowego internetu o wysokich przepustowościach
		Program Operacyjny Polska Cyfrowa na lata 2014-2020"
cell_b4 = "I. INFORMACJE O PROJEKCIE"
cell_b5 = "Numer wniosku o dofinansowanie"
cell_b6 = "Nazwa Wnioskodawcy"
cell_b7 = "Tytuł projektu"
cell_b8 = "Lokalizacja projektu - numer obszaru"
cell_b9 = "Kwota wydatków kwalifikowalnych"
cell_b10 = "Wnioskowana kwota dofinansowania"
cell_b11 = "Udział wnioskowanego dofinansowania w kosztach kwalifikowalnych projektu"
cell_b13 = "Okres realizacji projektu"
cell_g9_10 = "PLN"
cell_g11 = "%"
cell_b15 = "W odpowiedzi na każde z poniższych kryteriów należy zaznaczyć odpowiednio TAK / NIE z pola rozwijanego (niespełnienie kryterium oznacza odrzucenie wniosku)"
cell_b17_b49_b59_b68_b76 = "Lp."
cell_c17 = "II. KRYTERIA MERYTORYCZNE OCENA: TAK/NIE"
cell_g17 = "TAK/NIE"
cell_b18 = "2."
cell_c18 = "Obszar inwestycji"
cell_b19_b25_b31_b37_b43 = "Uzasadnienie"
cell_b24 = "3."
cell_c24 = "Wskaźniki produktu i rezultatu są:
		• adekwatne dla danego rodzaju projektu
		• realne do osiągnięcia
		• odzwierciedlają założone cele projektu"
cell_b30 = "4."
cell_c30 = "Koncepcja techniczna projektu jest zgodna z wymaganiami dla sieci NGA - POPC, oraz wymaganiami dla podłączenia szkół/placówek edukacyjnych"
cell_b36 = "5."
cell_c36 = "Harmonogram zadań projektu i kamieni milowych oraz zakres finansowy jest:
			• wykonalny/możliwy do przeprowadzenia,
			• uwzględnia czas niezbędny na przeprowadzenie procedur konkurencyjnego wyboru i wpływ czynników zewnętrznych"
cell_b42 = "8."
cell_c42 = "Efektywność realizacji projektu - ocena techniczna i koszty"



require 'rubyXL'
workbook = RubyXL::Workbook.new
workbook.write("form.xlsx")

workbook = RubyXL::Parser.parse("form.xlsx")
sheet = workbook["Sheet1"]
sheet.sheet_name = "1.1 POPC" #rename Sheet1 to 1.1 POPC

font_size = 11
font_name = 'Trebuchet MS'
column_number = 1
while column_number < 7
	sheet.change_column_font_name(column_number, font_name) #sets given column font to Trebuchet MS
	sheet.change_column_font_size(column_number, font_size) #sets given column font size to 11
	column_number += 1
end

column_width = [
	[0, 3.38],
	[1, 3.38],
	[2, 25.38],
	[3, 31.38],
	[4, 8.5],
	[5, 9.38],
	[6, 9.13]
]

#column width setter - values taken from array column_width
xx = 0

while xx < column_width.length
	a_cw = column_width[xx][0];
	b_cw = column_width[xx][1];
	sheet.change_column_width(a_cw, b_cw)
	xx += 1
end

workbook.write("./form.xlsx")

num_a = 1
num_b = 6

num_x = 3
num_y = 93

for y_y in num_a..num_b
  for x_x in num_x..num_y
	  c = workbook[0].add_cell(x_x, y_y)
	  c.change_border(:bottom, 'thin')
		c.change_border(:top, 'thin')
		c.change_border(:right, 'thin')
		c.change_border(:left, 'thin')
	end
end

workbook.write("./form.xlsx")

no_border_rows = [13, 14, 15, 47, 56, 57, 65, 74, 82, 86, 90]

num_a = 0
num_b = 6

for y_y in num_a..num_b
	no_border_rows.each do |row|
		c = workbook[0].add_cell(row, y_y)
	  c.change_border(:bottom, 'hairline')
		c.change_border(:top, 'hairline')
		c.change_border(:right, 'hairline')
		c.change_border(:left, 'hairline')
	end
end

workbook.write("./form.xlsx")

#merge columns
cells_to_merge = [[0, 1, 0, 6], [1, 1, 1, 6], [3, 1, 3, 6], [4, 1, 4, 2], [4, 3, 4, 6], [5, 1, 5, 2], [5, 3, 5, 6],
							[6, 1, 6, 2], [6, 3, 6, 6], [7, 1, 7, 2], [7, 3, 7, 6], [8, 1, 8, 2], [8, 3, 8, 5], [9, 1, 9, 2],
							[9, 3, 9, 5], [12, 1, 12, 2], [12, 3, 12, 6], [14, 1, 14, 6], [16, 2, 16, 5], [17, 2, 17, 5],
							[18, 1, 18, 6], [23, 2, 23, 5], [24, 1, 24, 6], [29, 2, 29, 5], [30, 1, 30, 6], [35, 2, 35, 5],
							[36, 1, 36, 6], [41, 2, 41, 5], [42, 1, 42, 6], [48, 2, 48, 5], [49, 2, 49, 5], [55, 2, 55, 5],
							[58, 2, 58, 5],	[59, 1, 60, 1], [59, 2, 60, 6], [61, 1, 62, 1], [61, 2, 62, 6],
							[63, 1, 64, 1], [63, 2, 64, 6], [65, 2, 65, 6],	[66, 2, 66, 4], [66, 5, 66, 6], [67, 1, 67, 6], [68, 1, 68, 6],
							[69, 1, 69, 6], [70, 1, 70, 6],
							[71, 1, 71, 6], [73, 1, 73, 2], [73, 3, 73, 6], [74, 1, 74, 2], [74, 3, 74, 6], [75, 1, 75, 2],
							[75, 3, 75, 6], [77, 1, 77, 2], [77, 3, 77, 6], [78, 1, 78, 2], [78, 3, 78, 6], [79, 1, 79, 2],
							[79, 3, 79, 6], [80, 1, 80, 2], [80, 3, 80, 6], [81, 1, 81, 2], [81, 3, 81, 6], [82, 1, 82, 2],
							[82, 3, 82, 6]]

	  
ctm1 = 0

while ctm1 < cells_to_merge.length
	a_cw = cells_to_merge[ctm1][0];
	b_cw = cells_to_merge[ctm1][1];
	c_cw = cells_to_merge[ctm1][2];
	d_cw = cells_to_merge[ctm1][3];
	sheet.merge_cells(a_cw, b_cw, c_cw, d_cw)
	ctm1 += 1
end

workbook.write("./form.xlsx")


#set rows height - Sets first row height to 27.75
rows_height = [
		[0, 27.75], [1, 129], [2, 18], [3, 18], [4, 123.5], [5, 57], [6, 57], [7, 18],
		[8, 18],[9, 18], [10, 24], [11, 24], [12, 18], [13, 18], [14, 36], [15, 18],
		[16, 18], [17, 18], [18, 18], [19, 18], [20, 18], [21, 18], [22, 18], [23, 72],
		[24, 18], [25, 18], [26, 18], [27, 18], [28, 18], [29, 36], [30, 18], [31, 18],
		[32, 18], [33, 18], [34, 18], [35, 72], [36, 18], [37, 18], [38, 18], [39, 18],
		[40, 18], [41, 18], [42, 18], [43, 18], [44, 18], [45, 18], [46, 18], [47, 18],
		[48, 18], [49, 0], [50, 0], [51, 0], [52, 0], [53, 0], [54, 0], [55, 57], [56, 0],
		[57, 18], [58, 18], [59, 18], [60, 18], [61, 18], [62, 18], [63, 18], [64, 18], 
		[65, 18], [66, 18], [67, 36], [68, 36], [69, 18], [70, 18], [71, 18], [72, 18], 
		[73, 18], [74, 18], [75, 18], [76, 36], [77, 18], [78, 18], [79, 18], [80, 18], 
		[81, 18], [82, 18], [83, 36], [84, 18], [85, 36], [86, 18], [87, 36], [88, 18], 
		[89, 36], [90, 18], [91, 36], [92, 18], [93, 36]]


rh1 = 0

while rh1 < rows_height.length
	a_rh = rows_height[rh1][0];
	b_rh = rows_height[rh1][1];
	sheet.change_row_height(a_rh, b_rh)
	rh1 += 1
end

workbook.write("./form.xlsx")

#merge rows
sheet.merge_cells(10, 1, 11, 2)  # Merges B10:C11
sheet.merge_cells(10, 3, 11, 5)  # Merges D10:F11
sheet.merge_cells(10, 6, 11, 6)  
sheet.merge_cells(19, 1, 22, 6)
sheet.merge_cells(25, 1, 28, 6)
sheet.merge_cells(31, 1, 34, 6)
sheet.merge_cells(37, 1, 40, 6)
sheet.merge_cells(43, 1, 46, 6)

workbook.write("./form.xlsx")

c = workbook[0].add_cell(0, 1, "#{cell_b1}")
c.change_font_bold(true)
c.change_font_size(10)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')

c = workbook[0].add_cell(1, 1, "#{cell_b2}")
c.change_font_bold(true)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_text_wrap(true)

cell_format_and_value = [
													[3, 1, cell_b4, true, 'd9d9d9', 'center', 'center', 11, false],
													[4, 1, cell_b5, false, 'ffffff', 'left', 'center', 10, false],
													[5, 1, cell_b6, false, 'ffffff', 'left', 'center', 10, false],
													[6, 1, cell_b7, false, 'ffffff', 'left', 'center', 10, false],
													[7, 1, cell_b8, false, 'ffffff', 'left', 'center', 10, false],
													[8, 1, cell_b9, false, 'ffffff', 'left', 'center', 10, false],
													[9, 1, cell_b10, false, 'ffffff', 'left', 'center', 10, false],
													[10, 1, cell_b11, false, 'ffffff', 'left', 'center', 10, true],
													[11, 1, '', false, 'ffffff', 'left', 'center', 10, true],
													[12, 1, cell_b13, false, 'ffffff', 'left', 'center', 10, false],
													[8, 6, cell_g9_10, false, 'ffffff', 'center', 'center', 10, false],
													[9, 6, cell_g9_10, false, 'ffffff', 'center', 'center', 10, false],
													[10, 6, cell_g11, false, 'ffffff', 'center', 'center', 10, true],
													[11, 6, '', false, 'ffffff', 'center', 'center', 10, false],
													[14, 1, cell_b15, false, 'ffffff', 'center', 'center', 11, true],
													[16, 1, cell_b17_b49_b59_b68_b76, true, 'd9d9d9', 'center', 'center', 11, false],
													[16, 2, cell_c17, true, 'd9d9d9', 'center', 'center', 11, false],
													[16, 6, cell_g17, true, 'd9d9d9', 'center', 'center', 11, false],
													[17, 1, cell_b18, false, 'ffffff', 'center', 'center', 11, false],
													[17, 2, cell_c18, false, 'ffffff', 'left', 'center', 11, false],
													[17, 6, '', true, 'ffffff', 'center', 'center', 11, false],
													[18, 1, cell_b19_b25_b31_b37_b43, true, 'ffffff', 'center', 'center', 11, false],
													[24, 1, cell_b19_b25_b31_b37_b43, true, 'ffffff', 'center', 'center', 11, false],
													[30, 1, cell_b19_b25_b31_b37_b43, true, 'ffffff', 'center', 'center', 11, false],
													[36, 1, cell_b19_b25_b31_b37_b43, true, 'ffffff', 'center', 'center', 11, false],
													[42, 1, cell_b19_b25_b31_b37_b43, true, 'ffffff', 'center', 'center', 11, false],
													[23, 1, cell_b24, false, 'ffffff', 'center', 'center', 11, false],
													[23, 2, cell_c24, false, 'ffffff', 'left', 'center', 11, true],
													[23, 6, '', true, 'ffffff', 'center', 'center', 11, false],
													[29, 1, cell_b30, false, 'ffffff', 'center', 'center', 11, false],
													[29, 2, cell_c30, false, 'ffffff', 'left', 'center', 11, true],
													[29, 6, '', true, 'ffffff', 'center', 'center', 11, false],
													[35, 1, cell_b36, false, 'ffffff', 'center', 'center', 11, false],
													[35, 2, cell_c36, false, 'ffffff', 'left', 'center', 11, true],
													[35, 6, '', true, 'ffffff', 'center', 'center', 11, false],
													[41, 1, cell_b42, false, 'ffffff', 'center', 'center', 11, false],
													[41, 2, cell_c42, false, 'ffffff', 'left', 'center', 11, false],
													[41, 6, '', true, 'ffffff', 'center', 'center', 11, false]
												]


cfav = 0

while cfav < cell_format_and_value.length
	a_cfav = cell_format_and_value[cfav][0];
	b_cfav = cell_format_and_value[cfav][1];
	c_cfav = cell_format_and_value[cfav][2];
	d_cfav = cell_format_and_value[cfav][3];
	e_cfav = cell_format_and_value[cfav][4];
	f_cfav = cell_format_and_value[cfav][5];
	g_cfav = cell_format_and_value[cfav][6];
	h_cfav = cell_format_and_value[cfav][7];
	i_cfav = cell_format_and_value[cfav][8];
		c = workbook[0][a_cfav][b_cfav]
		c.raw_value = c_cfav
		c.change_font_bold(d_cfav)
		c.change_fill(e_cfav)
		c.change_horizontal_alignment(f_cfav)
		c.change_vertical_alignment(g_cfav)
		c.change_font_size(h_cfav)
		c.change_text_wrap(i_cfav)
	cfav += 1
end


c = workbook[0][14][1]
c.change_font_color('ff0000')


workbook.write("./form.xlsx")