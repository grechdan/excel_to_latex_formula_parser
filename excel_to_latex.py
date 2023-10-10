import openpyxl
import re
from pylatex import Document, Section, Math, Alignat, Package

predefined_values = {'Beton':'Beton','f_ck':'f ck','f_cd':'f cd','f_cm':'f cm','f_ctm':'f ctm','f_ctk0.05':'f ctk0.05','E_cm':'E cm','ν_':"ν'",'ν':'ν','f_yk':'f yk','f_yd':'f yd','E_s':'E s',
'd_g':'d g','cotθ':'cotθ','L_d':'L d','L_t':'L t','h_d':'h d','h_t':'h t','b_t':'b t','b_s':'b s','h_s':'h s','c_':'c'}
greek_letters = {'α':'\\alpha{}','β':'\\beta{}','γ':'\gamma{}','δ':'\delta{}','ε':'\epsilon{}','ζ':'\zeta{}','η':'\eta{}','θ':'\\theta{}','ι':'\iota{}','κ':'\kappa{}','λ':'\lambda{}','μ':'\mu{}','ν':'\\nu{}',
'ξ':'\\xi{}','ο':'o{}','π':'\pi{}','ρ':'\\rho{}','σ':'\sigma{}','τ':'\tau{}','υ':'\\upsilon{}','ϕ':'\\varphi{}','χ':'\chi{}','ψ':'\psi{}','ω':'\omega{}','Δ':'\Delta{}','Σ':'\Sigma{}','⌀':'\diameter{}'}


class Line():

    def __init__(self, name, latex_name):

        self.name = name
        self.latex_name = latex_name
        self.row_number = ''
        self.formula_excel = ''
        self.substituted_formula = ''
        self.substituted_values = ''
        self.value = None
        self.units = ''

class Cell:

    def __init__(self, coordinate: str, formula_excel: str, formula_value: str):
        
        self.coordinate = coordinate
        self.value = self.value_classificator(formula_excel, formula_value) # None/Value/Link/Formula
        self.value_type = '' # 'val'/'link'/'form'/'text' 

    def value_classificator(self, formula_excel: str, value: str):

        if formula_excel == value:

            try:

                return Value(float(value))
            
            except ValueError:

                return Text(value)

        else:

            match = re.fullmatch(r'(^[B]+\d+$)', formula_excel)
            
            if match != None:
                
                return Link(float(value), formula_excel)
            
            else:
                
                return Formula(float(value), formula_excel)

class Value:

    def __init__(self, value: float):

        self.float_value = value

class Text():
    
    def __init__(self, value: str):

        self.str_value = value

class Link:

    def __init__(self, value: float, cell_coordinate: str):
        
        self.to_cell = cell_coordinate
        self.float_value = value   

class Formula():

    def __init__(self, value: float, formula_excel: str):

        self.float_value = value 
        self.excel = formula_excel
        self.left_side = None
        self.operator = None
        self.right_side = None
        self.latex_with_values = self.formula_parcer()
        #print(self.latex_with_values)
        self.latex_with_formulas = ''

    def power(self, splitted: list):

        base_type = self.elements_classificator(splitted[0])
        power_type = self.elements_classificator(splitted[1][:-1])
        string = ''
        base = ''
        power = ''

        if base_type == 0:
            
            base = str(splitted[0])
            
            if  power_type == 0:

                power = str(splitted[1][:-1])
        
            elif power_type == 1:

                power = str(splitted[1][:-1])
        
            elif power_type == 2:

                power = Formula(0, splitted[1][:-1]).latex_with_values
        
        elif base_type == 1:

            base = str(splitted[0])
            
            if  power_type == 0:

                power = str(splitted[1][:-1])
        
            elif power_type == 1:

                power = str(splitted[1][:-1])
        
            elif power_type == 2:

                power = Formula(0, splitted[1][:-1]).latex_with_values
        
        elif base_type == 2:

            base = Formula(0, splitted[0]).latex_with_values
            
            if  power_type == 0:

                power = str(splitted[1][:-1])
        
            elif power_type == 1:

                power = str(splitted[1][:-1])
        
            elif power_type == 2:

                power = Formula(0, splitted[1][:-1]).latex_with_values
        
        return string

    def excel_operations(self) -> str:

        string = ''

        if self.left_side == 'POWER':

            base, power = self.two_elements_operations(self.right_side.split(','))
            string = f'({base} ^ {power})'

        if self.left_side == 'VALUE':

            pass

        if self.left_side == 'MID':

            pass

        return string
    
    def round_brackets(self):

        print('-------------------')

        print(self.excel)
        string = ''

    
        match = re.search(r'\((.*)\)', self.excel)

        inside = match.group(0)
        
        print(self.excel.split(inside))

    
        return string

    def two_elements_operations(self) -> tuple:

        first_element_type = self.elements_classificator(self.left_side)
        second_element_type = self.elements_classificator(self.right_side)
        print('------------------------------')
        print(self.excel, '---',self.left_side, '---', self.right_side)
        print(first_element_type, second_element_type)
        first_element = ''
        second_element = ''

        if first_element_type == 0:
            
            first_element = str(self.left_side)
            
            if  second_element_type == 0:

                second_element = str(self.right_side)
        
            elif second_element_type == 1:

                second_element = str(self.right_side)
        
            elif second_element_type == 2:

                second_element = Formula(0, self.right_side).latex_with_values
        
        elif first_element_type == 1:

            first_element = str(self.left_side)
            
            if  second_element_type == 0:

                second_element = str(self.right_side)
        
            elif second_element_type == 1:

                second_element = str(self.right_side)
        
            elif second_element_type == 2:

                second_element = Formula(0, self.right_side).latex_with_values
        
        elif first_element_type == 2:

            first_element = Formula(0, self.left_side).latex_with_values
            
            if  second_element_type == 0:

                second_element = str(self.right_side)
        
            elif second_element_type == 1:

                second_element = str(self.right_side)
        
            elif second_element_type == 2:

                second_element = Formula(0, self.right_side).latex_with_values

        return first_element, second_element
    
    def basic_operations(self) ->str:

        if self.operator == '+':

            augend, addend = self.two_elements_operations()
            string = f'({augend} + {addend})'

        elif self.operator == '-':

            minuend, subtrahend = self.two_elements_operations()
            string = f'({minuend} - {subtrahend})'

        elif self.operator == '*':

            multiplier, multiplicand = self.two_elements_operations()
            string = f'({multiplier} * {multiplicand})'

        elif self.operator == '/':

            dividend, divisor = self.two_elements_operations()
            string = f'({dividend} / {divisor})'

        elif self.operator == '^':

            base, power = self.two_elements_operations()
            string = f'({base} ^ {power})'
    
        return string
    
    def formula_parcer(self) -> str:

        string = ''
        #r'([A-Za-z0-9.]+)?\s*([\*\-\+\/])?\s*([A-Z]+)\((.*)\)'
        #r'^([A-Z]+\(.*?\)|[A-Z0-9]+)([\*\/+-])(.*)$'
        #r'^(.*\))(?:([\*\/+-])(.*))?$'
        #r'^(.*\([^()]*\))(?:([\*\/+-])(.*))?$'
        excel_formula_match_efo = re.compile(r'^(.*\))(?:([\*\/+-])(.*))?$').match(self.excel)# search for excel function and operation
        #excel_formula_match1 = re.compile(r'([A-Za-z0-9.]+)?\s*([\*\-\+\/])?\s*([A-Z]+)\((.*)\)').match(self.excel)
        excel_formula_match_as = re.compile(r'^(.*?)([+-])(.*)$').match(self.excel)# search for addition and subtraction
        excel_formula_match_mdp = re.compile(r'^(.*?)([*\/\^])(.*)$').match(self.excel)# search for multiplication, division, and power
        #excel_formula_match3 = re.compile(r'^(.*?)([*\/+\-])(.*)$').match(self.excel)

        excel_formula_match_efc = re.compile(r'([A-Z]+)\((.*)\)').match(self.excel)# search for excel function and content

        #round_brackets_match1 = re.compile(r'\((.*)\)').match(self.excel)
        flag = 0
        
        if excel_formula_match_as:
            
            self.left_side, self.operator, self.right_side = excel_formula_match_as.groups()   

            if self.left_side != None and self.left_side.count('(') == self.left_side.count(')'):
                
                flag = 1
                
                print(f"excel_formula_match_as: {self.left_side}, {self.operator}, {self.right_side}")
                self.basic_operations()
        
        if excel_formula_match_mdp and flag == 0:
            
            self.left_side, self.operator, self.right_side = excel_formula_match_mdp.groups()

            if self.left_side != None and self.left_side.count('(') == self.left_side.count(')'):

                flag = 1
                print(f"excel_formula_match_md: {self.left_side}, {self.operator}, {self.right_side}")
                self.basic_operations()

        if excel_formula_match_efo and flag == 0:
            
            self.left_side, self.operator, self.right_side = excel_formula_match_efo.groups()

            if self.operator != None and self.right_side != None:
                
                flag = 1

                print(f"excel_formula_match: {self.left_side}, {self.operator}, {self.right_side}")
                self.basic_operations()

        if excel_formula_match_efc and flag == 0:
            
            self.left_side, self.right_side = excel_formula_match_efc.groups()

            print(f"excel_formula_match2: {self.left_side}, {self.operator}, {self.right_side}")
            #self.excel_operations()

        #if excel_formula_match and ((operator == None and right_side == None) or ('(' in left_side and ')' not in left_side)):
            
            #before, operator, function_name, function_args = excel_formula_match1.groups()
            #print(f"excel_formula_match1: {excel_formula_match.groups()}")

        #if excel_formula_match2 and (operator == None and right_side == None):

            #print(f"excel_formula_match2: {excel_formula_match2.group()}")

        """elif '+' in self.excel:

            augend, addend = self.two_elements_operations(self.excel.split('+', 1))
            string = f'({augend} + {addend})'

        elif '-' in self.excel:

            minuend, subtrahend = self.two_elements_operations(self.excel.split('-', 1))
            string = f'({minuend} - {subtrahend})'

        elif '*' in self.excel:

            multiplier, multiplicand = self.two_elements_operations(self.excel.split('*', 1))
            string = f'({multiplier} * {multiplicand})'

        elif '/' in self.excel:

            dividend, divisor = self.two_elements_operations(self.excel.split('/', 1))
            string = f'({dividend} / {divisor})'

        elif '^' in self.excel:

            base, power = self.two_elements_operations(self.excel.split('^', 1))
            string = f'({base} ^ {power})'"""

        return string

    def elements_classificator(self, formula_excel: str) -> int:

        try:

            float(formula_excel)
            return 0
        
        except ValueError:

            match = re.fullmatch(r'(^[B]+\d+$)', formula_excel)
            
            if match != None:

                return 1
            
            else:

                return 2

                """if ' ' in formula_excel:

                    return -1

                if formula_excel != None and formula_excel != 'None':

                    formula_parced = self.formula_parcer()
                    
                    return formula_excel"""


# Load the Excel file
workbook = openpyxl.load_workbook('H.xlsx',data_only=True)

# Select a specific sheet
sheet = workbook.active

#Replace the predefined variable with a general name
def predefined_replacer(value):
    for cell in sheet['A']:
        if cell.value == value:
            value = "B"+str(cell.row)
    return value

def fraction(formula):
    
    dividend_pattern = r'([B]+\d+$)|(\d+(\.\d*)?$)'
    divisor_pattern = r'(^[B]+\d+)|(^\d+(\.\d*)?)'

    dividend, divisor = formula.split('/')
    dividend = re.search(dividend_pattern,dividend).group()
    divisor = re.search(divisor_pattern,divisor).group()
    fraction = dividend + '/' + divisor

    if dividend[-1] == ')':
        print(dividend[-1])
    else:
        print(dividend)
        latex_dividend = '\\frac{' + dividend + '}'
        latex_fraction = fraction.replace(dividend,latex_dividend)

    if divisor[0] == '(':
        print(divisor[0])
    else:
        latex_divisor = '{' + divisor + '}'
        latex_fraction.replace(divisor,latex_divisor)
    
    return formula.replace(fraction,latex_fraction)

# Getting the formula from the cell
def formula(coordinate):
    
    workbook = openpyxl.load_workbook('H.xlsx',data_only=False)
    sheet = workbook.active
    cell = sheet[coordinate]
    formula = str(cell.value).split('=',1)[-1]
    for k in predefined_values:
        if k in formula:
            formula = formula.replace(k,predefined_replacer(predefined_values[k]))

    return formula

# Round function for the result values
def rounder(cell):
    cell = str(cell.value)
    if '.' in cell:   
        decimal_part = cell.split('.')[-1]
        if len(decimal_part)>=3:  
            cell = round(float(cell),3)
        return str(cell)
    else:
        return str(cell)
 
# Splitting the first cell on vataible and it's atribute
def subscript(coordinate):
    
    cell=sheet[coordinate]
    cell_check=sheet["C"+str(cell.row)]
    latex_cell=cell.value

    if cell_check.value != None:
        for k in greek_letters:
            if k in latex_cell:
                latex_cell = latex_cell.replace(k,greek_letters[k])

            else:
                latex_cell
        
        if ' ' in latex_cell:
            var_atr_list = latex_cell.split(' ')
            var_atr_list[-1] = '{'+var_atr_list[-1]+'}'
            return '_'.join(var_atr_list)

        else:
            return latex_cell
    
    else:
        return latex_cell

#Replacing cell's coordinate with value
def value_replacer(cell):
    
    pattern_cell = r'[B]+\d+'

    formula_general = formula(cell.coordinate)
    formula_list = re.findall(pattern_cell, formula_general)

    for i in formula_list:
        if len(formula_list) == 0:
            break
        c=sheet[i]
        formula_general = formula_general.replace(i,rounder(c))

    return formula_general   

#Replacing cell's coordinate with attribute
def formula_replacer(cell):
    
    pattern_cell = r'[B]+\d+'

    formula_general = formula(cell.coordinate)
    formula_list = re.findall(pattern_cell, formula_general)

    for i in formula_list:
        if len(formula_list) == 0:
            break
        attribute = i.replace('B','A')
        c = sheet[attribute]
        formula_general = formula_general.replace(i,subscript(c.coordinate))

    return formula_general

"""def excel_formula(cell):

    pattern_function = r'\w+\(.+\)'
    functions = re.findall(pattern_function, cell)
    return functions"""

def alignat_fill(line_list, filling_string, j):

    for i in range(len(line_list)):
            

            if i == 0:
                
                filling_string += line_list[i] + '&='

            elif i == len(line_list) - 1:

                filling_string += line_list[i]
            
            else:

                filling_string += line_list[i] + '='

    filling_string += r'\\'
    

    return filling_string

lines = []

for row in sheet:
    
    for cell in row:
            
        if cell.column_letter == 'A':
            
            lines.append(Line(cell.value,subscript(cell.coordinate)))

        elif cell.column_letter == 'B':
            
            lines[-1].row_number = cell.coordinate
            lines[-1].formula_excel = formula(cell.coordinate)
            lines[-1].substituted_formula = formula_replacer(cell)
            lines[-1].substituted_values = value_replacer(cell)
            lines[-1].value = cell.value
            """print(formula_replacer(cell))
            print(value_replacer(cell))
            print(rounder(cell))
            """
        
        elif cell.column_letter == 'C':
            lines[-1].units = cell.value

for line in lines:

    if line.value != '8x⌀12A s =905':
        
        print('================================================')
        print(line.formula_excel)
        cell = Cell(line.row_number,line.formula_excel, str(line.value))

    
"""
        self.name = name
        self.latex_name = latex_name
        self.row_number = row_number
        self.formula_excel = ''
        self.substituted_formula = ''
        self.substituted_values = ''
        self.value = 0
        self.units = ''
cell = sheet['B4']
print(fraction(formula(cell.coordinate)))
"""
"""lin = [item for item in lin if 'ArrayFormula' not in item]
    lin = [item for index, item in enumerate(lin) if item not in lin[:index]]
    if lin[0] == 'Návrh':
        lin = [item for index, item in enumerate(lin) if index not in (1,2) ]
        lin[1] = lin[1].replace('⌀',greek_letters['⌀'])

    if len(lin) > 1:
        lin[-2] += lin[-1]
        del lin[-1]
    all_lin.append(lin)"""
