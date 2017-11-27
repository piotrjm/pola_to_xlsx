system "cls"

a = "Załącznik nr 2c - Wzór formularza oceny - ocena merytoryczna II stopnia"
bb = "II stopnia"
b = "FORMULARZ OCENY - OCENA MERYTORYCZNA
		Centrum Projektów Polska Cyfrowa
		w ramach oceny #{bb}
		I Oś priorytetowa – Powszechny dostęp do szybkiego internetu
		Działanie 1.1 – Wyeliminowanie terytorialnych różnic w możliwości dostępu do szerokopasmowego internetu o wysokich przepustowościach
		Program Operacyjny Polska Cyfrowa na lata 2014-2020"
cc = "I. INFORMACJE O PROJEKCIE"
d = "Numer wniosku o dofinansowanie"
e = "Nazwa Wnioskodawcy"
f = "Tytuł projektu"
g = "Lokalizacja projektu - numer obszaru"
h = "Kwota wydatków kwalifikowalnych"
i = "Wnioskowana kwota dofinansowania"
j = "Udział wnioskowanego dofinansowania w kosztach kwalifikowalnych projektu"
k = "Okres realizacji projektu"
l = "PLN"
m = "%"
n = "W odpowiedzi na każde z poniższych kryteriów należy zaznaczyć odpowiednio TAK / NIE z pola rozwijanego (niespełnienie kryterium oznacza odrzucenie wniosku)"
o = "Lp."
p = "II. KRYTERIA MERYTORYCZNE OCENA: TAK/NIE"
q = "TAK/NIE"
r = "2."
s = "Obszar inwestycji"
t = "NIE"
u = "Uzasadnienie"
v = "3."
w = "Wskaźniki produktu i rezultatu są:
		• adekwatne dla danego rodzaju projektu
		• realne do osiągnięcia
		• odzwierciedlają założone cele projektu"
x = "4."
y = "Koncepcja techniczna projektu jest zgodna z wymaganiami dla sieci NGA - POPC, oraz wymaganiami dla podłączenia szkół/placówek edukacyjnych"
z = "5."
za = "Harmonogram zadań projektu i kamieni milowych oraz zakres finansowy jest:
			• wykonalny/możliwy do przeprowadzenia, 
			• uwzględnia czas niezbędny na przeprowadzenie procedur konkurencyjnego wyboru i wpływ czynników zewnętrznych"
zb = "8."
zc = "Efektywność realizacji projektu - ocena techniczna i koszty"



require 'rubyXL'
workbook = RubyXL::Workbook.new
workbook.write("form.xlsx")

workbook = RubyXL::Parser.parse("form.xlsx")
sheet = workbook["Sheet1"]
sheet.sheet_name = "1.1 POPC" #rename Sheet1 to Test1
workbook.write("./form.xlsx")



sheet = workbook["1.1 POPC"]
sheet.change_column_font_name(1, 'Trebuchet MS')   # Makes first column have font Courier
sheet.change_column_font_name(2, 'Trebuchet MS') 
sheet.change_column_font_name(3, 'Trebuchet MS') 
sheet.change_column_font_name(4, 'Trebuchet MS') 
sheet.change_column_font_name(5, 'Trebuchet MS') 
sheet.change_column_font_name(6, 'Trebuchet MS') 

# sheet.add_cell(3, 0,'test')  # Sets cell A4 to string "test"
# sheet.sheet_data[3][0].change_horizontal_alignment('center') # Sets cell A4 to center horizontal aligment
# sheet.sheet_data[3][0].change_border(:bottom, 'thin')

# c = workbook[0].add_cell(6,0)
# c.set_number_format('yyyy/mm/dd')
# c.change_contents(Date.today)
# c.change_fill('D9D9D')
# c.change_font_bold(true)
# c.change_font_italics(true)
# c.change_font_size(20)
# c.change_font_color('e6c3b6')

sheet.change_column_width(0, 3.38)  # Sets first column width to 3.38
sheet.change_column_width(1, 3.38)  # Sets second column width to 3.38
sheet.change_column_width(2, 25.38)  # Sets third column width to 20
sheet.change_column_width(3, 31.38)  # Sets fourth column width to 20
sheet.change_column_width(4, 8.5)  # Sets fifth column width to 20
sheet.change_column_width(5, 9.38)  # Sets sixth column width to 20
sheet.change_column_width(6, 9.13)  # Sets seventyh column width to 20
workbook.write("./form.xlsx")

#merge columns - head rows color 'D9D9D9'
sheet.merge_cells(0, 1, 0, 6)  # Merges B1:G1
sheet.merge_cells(1, 1, 1, 6)  # Merges B2:G2
sheet.merge_cells(3, 1, 3, 6)  # Merges B4:G4
sheet.merge_cells(4, 1, 4, 2)  # Merges B5:C5
sheet.merge_cells(4, 3, 4, 6)  # Merges D5:G5
sheet.merge_cells(5, 1, 5, 2)  
sheet.merge_cells(5, 3, 5, 6)  
sheet.merge_cells(6, 1, 6, 2)  
sheet.merge_cells(6, 3, 6, 6)  
sheet.merge_cells(7, 1, 7, 2)  
sheet.merge_cells(7, 3, 7, 6)  
sheet.merge_cells(8, 1, 8, 2)  
sheet.merge_cells(8, 3, 8, 5) 
sheet.merge_cells(9, 1, 9, 2)  
sheet.merge_cells(9, 3, 9, 5) 
# sheet.merge_cells(10, 1, 10, 2)  
# sheet.merge_cells(10, 3, 10, 5) 
# sheet.merge_cells(11, 1, 11, 2)  
# sheet.merge_cells(11, 3, 11, 5) 
sheet.merge_cells(12, 1, 12, 2)  
sheet.merge_cells(12, 3, 12, 6) 
sheet.merge_cells(14, 1, 14, 6)
sheet.merge_cells(16, 2, 16, 5)
sheet.merge_cells(17, 2, 17, 5)
sheet.merge_cells(18, 1, 18, 6)
#sheet.merge_cells(19, 1, 19, 6)
#sheet.merge_cells(20, 1, 20, 6)
#sheet.merge_cells(21, 1, 21, 6)
#sheet.merge_cells(22, 1, 22, 6)
sheet.merge_cells(23, 2, 23, 5)
sheet.merge_cells(24, 1, 24, 6)
# sheet.merge_cells(25, 1, 25, 6)
# sheet.merge_cells(26, 1, 26, 6)
# sheet.merge_cells(27, 1, 27, 6)
# sheet.merge_cells(28, 1, 28, 6)
sheet.merge_cells(29, 2, 29, 5)
sheet.merge_cells(30, 1, 30, 6)
# sheet.merge_cells(31, 1, 31, 6)
# sheet.merge_cells(32, 1, 32, 6)
# sheet.merge_cells(33, 1, 33, 6)
# sheet.merge_cells(34, 1, 34, 6)
sheet.merge_cells(35, 2, 35, 5)
sheet.merge_cells(36, 1, 36, 6)
# sheet.merge_cells(37, 1, 37, 6)
# sheet.merge_cells(38, 1, 38, 6)
# sheet.merge_cells(39, 1, 39, 6)
# sheet.merge_cells(40, 1, 40, 6)
sheet.merge_cells(41, 2, 41, 5)
sheet.merge_cells(42, 1, 42, 6)
# sheet.merge_cells(43, 1, 43, 6)
# sheet.merge_cells(44, 1, 44, 6)
# sheet.merge_cells(45, 1, 45, 6)
# sheet.merge_cells(46, 1, 46, 6)

sheet.merge_cells(48, 2, 48, 5)
sheet.merge_cells(49, 2, 49, 5)

sheet.merge_cells(51, 2, 51, 6)
sheet.merge_cells(52, 2, 52, 6)
sheet.merge_cells(53, 2, 53, 6)
sheet.merge_cells(54, 2, 54, 6)

sheet.merge_cells(56, 1, 56, 6)
sheet.merge_cells(57, 2, 57, 5)
sheet.merge_cells(58, 2, 58, 5)
sheet.merge_cells(59, 1, 59, 6)
sheet.merge_cells(60, 1, 60, 6)
sheet.merge_cells(61, 1, 61, 6)
sheet.merge_cells(62, 1, 62, 6)
sheet.merge_cells(63, 1, 63, 6)

#final decision
sheet.merge_cells(65, 2, 65, 6)
sheet.merge_cells(66, 2, 66, 4)
sheet.merge_cells(66, 5, 66, 6)
sheet.merge_cells(67, 1, 67, 6)
sheet.merge_cells(68, 1, 68, 6)
sheet.merge_cells(69, 1, 69, 6)
sheet.merge_cells(70, 1, 70, 6)
sheet.merge_cells(71, 1, 71, 6)

#first evaluator
sheet.merge_cells(73, 1, 73, 2)
sheet.merge_cells(73, 3, 73, 6)
sheet.merge_cells(74, 1, 74, 2)
sheet.merge_cells(74, 3, 74, 6)
sheet.merge_cells(75, 1, 75, 2)
sheet.merge_cells(75, 3, 75, 6)

#second evaluator
sheet.merge_cells(77, 1, 77, 2)
sheet.merge_cells(77, 3, 77, 6)
sheet.merge_cells(78, 1, 78, 2)
sheet.merge_cells(78, 3, 78, 6)
sheet.merge_cells(79, 1, 79, 2)
sheet.merge_cells(79, 3, 79, 6)

#third evaluator
sheet.merge_cells(80, 1, 80, 2)
sheet.merge_cells(80, 3, 80, 6)
sheet.merge_cells(81, 1, 81, 2)
sheet.merge_cells(81, 3, 81, 6)
sheet.merge_cells(82, 1, 82, 2)
sheet.merge_cells(82, 3, 82, 6)


workbook.write("./form.xlsx")

#set rows height
sheet.change_row_height(0, 27.75)  # Sets first row height to 27.75
sheet.change_row_height(1, 129)  # Sets second row height to 129
sheet.change_row_height(2, 18)  # Sets third row height to 18
sheet.change_row_height(3, 18)
sheet.change_row_height(4, 123.5)
sheet.change_row_height(5, 57)
sheet.change_row_height(6, 57)
sheet.change_row_height(7, 18)
sheet.change_row_height(8, 18)
sheet.change_row_height(9, 18)
sheet.change_row_height(10, 24)
sheet.change_row_height(11, 24)
sheet.change_row_height(12, 18)
sheet.change_row_height(13, 18)
sheet.change_row_height(14, 36)
sheet.change_row_height(15, 18)
sheet.change_row_height(16, 18)
sheet.change_row_height(17, 18)
sheet.change_row_height(18, 18)

sheet.change_row_height(19, 18)
sheet.change_row_height(20, 18)
sheet.change_row_height(21, 18)
sheet.change_row_height(22, 18)

sheet.change_row_height(23, 72)
sheet.change_row_height(24, 18)

sheet.change_row_height(25, 18)
sheet.change_row_height(26, 18)
sheet.change_row_height(27, 18)
sheet.change_row_height(28, 18)

sheet.change_row_height(29, 36)
sheet.change_row_height(30, 18)

sheet.change_row_height(31, 18)
sheet.change_row_height(32, 18)
sheet.change_row_height(33, 18)
sheet.change_row_height(34, 18)

sheet.change_row_height(35, 72)
sheet.change_row_height(36, 18)

sheet.change_row_height(37, 18)
sheet.change_row_height(38, 18)
sheet.change_row_height(39, 18)
sheet.change_row_height(40, 18)

sheet.change_row_height(41, 18)
sheet.change_row_height(42, 18)

sheet.change_row_height(43, 18)
sheet.change_row_height(44, 18)
sheet.change_row_height(45, 18)
sheet.change_row_height(46, 18)
sheet.change_row_height(47, 18)

sheet.change_row_height(48, 18)

sheet.change_row_height(49, 18)
sheet.change_row_height(50, 18)
sheet.change_row_height(51, 18)
sheet.change_row_height(52, 18)
sheet.change_row_height(53, 18)
sheet.change_row_height(54, 18)

sheet.change_row_height(55, 57)

sheet.change_row_height(56, 18)
sheet.change_row_height(57, 18)

sheet.change_row_height(58, 18)

sheet.change_row_height(59, 18)
sheet.change_row_height(60, 18)
sheet.change_row_height(61, 18)
sheet.change_row_height(62, 18)
sheet.change_row_height(63, 18)
sheet.change_row_height(64, 18)
sheet.change_row_height(65, 18)

sheet.change_row_height(66, 18)
sheet.change_row_height(67, 36)
sheet.change_row_height(68, 36)
sheet.change_row_height(69, 18)

sheet.change_row_height(70, 18)
sheet.change_row_height(71, 18)
sheet.change_row_height(72, 18)
sheet.change_row_height(73, 18)
sheet.change_row_height(74, 18)

sheet.change_row_height(75, 18)
sheet.change_row_height(76, 36)
sheet.change_row_height(77, 18)

sheet.change_row_height(78, 18)
sheet.change_row_height(79, 18)
sheet.change_row_height(80, 18)
sheet.change_row_height(81, 18)
sheet.change_row_height(82, 18)

sheet.change_row_height(83, 36)
sheet.change_row_height(84, 18)
sheet.change_row_height(85, 36)

sheet.change_row_height(86, 18)

sheet.change_row_height(87, 36)
sheet.change_row_height(88, 18)
sheet.change_row_height(89, 36)

sheet.change_row_height(90, 18)

sheet.change_row_height(91, 36)
sheet.change_row_height(92, 18)
sheet.change_row_height(93, 36)
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

c = workbook[0].add_cell(0, 1, "#{a}")
c.change_font_bold(true)
c.change_font_size(10)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(0, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(0, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(0, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(0, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(0, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(1, 1, "#{b}")
c.change_font_bold(true)
c.change_font_size(11)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_text_wrap(true)
c.change_border(:bottom, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(1, 2)
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(1, 3)
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(1, 4)
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(1, 5)
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(1, 6)
c.change_border(:bottom, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(3, 1, "#{cc}")
c.change_font_bold(true)
c.change_font_size(11)
c.change_fill('d9d9d9')
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(3, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(3, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(3, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(3, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(3, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(4, 1, "#{d}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(4, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(4, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(4, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(4, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(4, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(5, 1, "#{e}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(5, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(5, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(5, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(5, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(5, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(6, 1, "#{f}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(6, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(6, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(6, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(6, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(6, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(7, 1, "#{g}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(7, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(7, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(7, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(7, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(7, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(8, 1, "#{h}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(8, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(8, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(8, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(8, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(8, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(9, 1, "#{i}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(9, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(9, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(9, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(9, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(9, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(10, 1, "#{j}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_text_wrap(true)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(10, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(10, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(10, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(10, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(10, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(11, 1)
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_text_wrap(true)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(11, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(11, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(11, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(11, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(11, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(12, 1, "#{k}")
c.change_font_size(10)
c.change_horizontal_alignment('left')
c.change_vertical_alignment('center')
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:left, 'thin')

c = workbook[0].add_cell(12, 2)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(12, 3)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(12, 4)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(12, 5)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')

c = workbook[0].add_cell(12, 6)
c.change_border(:bottom, 'thin')
c.change_border(:top, 'thin')
c.change_border(:right, 'thin')

c = workbook[0].add_cell(8, 6, "#{l}")
c.change_font_size(10)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_border(:right, 'thin')
c.change_border(:left, 'thin')
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(9, 6, "#{l}")
c.change_font_size(10)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_border(:right, 'thin')
c.change_border(:left, 'thin')
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(10, 6, "#{m}")
c.change_font_size(10)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_border(:right, 'thin')
c.change_border(:left, 'thin')
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(11, 6)
c.change_font_size(10)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')
c.change_border(:right, 'thin')
c.change_border(:left, 'thin')
c.change_border(:bottom, 'thin')

c = workbook[0].add_cell(14, 1, "#{n}")
c.change_font_size(11)
c.change_font_color('ff0000')
c.change_text_wrap(true)
c.change_horizontal_alignment('center')
c.change_vertical_alignment('center')


workbook.write("./form.xlsx")