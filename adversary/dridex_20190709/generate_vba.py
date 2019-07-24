import string
import random
import os

vba_variables = dict()

english_wordlist_path = os.path.join('\\'.join(__file__.split("/")[0:-3]), "internals", "words.txt")

with open(english_wordlist_path, 'r') as r:
    words = set(r.readlines())

used_variable_names = []

def generate_variable_name():
    first_character = "a"
    number_of_additional_characters = random.randint(4, 8)
    additional_character_options = list(string.ascii_letters)
    for d in string.digits:
        additional_character_options.append(d)

    var_to_return = first_character

    var_length_counter = 0

    while var_length_counter < number_of_additional_characters:
        var_to_return = str(var_to_return + random.choice(additional_character_options))
        var_length_counter = var_length_counter + 1

    return var_to_return

def get_new_variable_name():
    while True:
        candidate = generate_variable_name()
        if candidate not in used_variable_names:
            used_variable_names.append(candidate)
            return candidate

def get_new_unused_variable_and_definition():
    this_var = get_new_variable_name()

    this_choice = random.randint(1, 4)

    if this_choice == 1:  # bool statement
        return "Dim {} As {}\r\n{} = {}\r\n".format(
            this_var,
            "Boolean",
            this_var,
            random.choice(["True", "False"])
        )
    elif this_choice == 2:  # int statement
        return "Dim {} As {}\r\n{} = {}\r\n".format(
            this_var,
            "Integer",
            this_var,
            str(random.randint(-32000, 32000))
        )
    elif this_choice == 3:  # int statement
        return "Dim {} As {}\r\n{} = {}\r\n".format(
            this_var,
            "Long",
            this_var,
            str(random.randint(-32000, 32000))
        )
    elif this_choice == 4:  # int statement
        return "Dim {} As {}\r\n{} = \"{}\"\r\n".format(
            this_var,
            "String",
            this_var,
            get_new_variable_name()
        )

def get_new_comment():
    this_comment_word_count = random.randint(0, 5)
    this_comment = "'"
    comment_counter = 0
    while comment_counter < this_comment_word_count:
        trying = 1
        while trying == 1:
            wordattempt = str((random.sample(words, k=1)[0].rstrip("\n")))
            if wordattempt.isalpha():
                if comment_counter == 0:
                    wordattempt = wordattempt.title()
                this_comment = this_comment + " " + wordattempt
                trying = 0
        comment_counter = comment_counter + 1
    this_comment = this_comment + "\r\n"
    return this_comment

def get_var_type():
    this_choice = random.randint(1, 4)
    if this_choice == 1:
        return "Boolean"
    elif this_choice == 2:
        return "Integer"
    elif this_choice == 3:
        return "Long"
    elif this_choice == 4:
        return "String"

#  ThisDocument
mod_1_name = "ThisDocument"
mod_1 = ""

mod_1 = mod_1 + "Sub Document_Open()\r\n"
for i in range(1, random.randint(8, 11)):  # between 8 and 11 sets of filler code
    mod_1 = mod_1 + get_new_unused_variable_and_definition()
    if random.randint(1, 7) < 7:  # 6 of 7 times, insert a comment after
        mod_1 = mod_1 + get_new_comment()
    if random.randint(1, 5) == 1:  # 1  of 5 times, insert a new line
        mod_1 = mod_1 + "\r\n"

mod_1 = mod_1 + "main\r\nEnd Sub"

#  Module 2 - this module contains no functional code
mod_2_name = get_new_variable_name()
mod_2 = ""

# in sample looks like:
# Function aoYbuI(aX2A6F9V As Long) As String
mod_2 = mod_2 + "Function {}({} As {}) As {}\r\n".format(
    get_new_variable_name(),
    get_new_variable_name(),
    get_var_type(),
    get_var_type()
)

for i in range(1, random.randint(10, 12)):
    mod_2 = mod_2 + get_new_unused_variable_and_definition()
    if random.randint(1, 7) < 7:  # 6 of 7 times, insert a comment after
        mod_2 = mod_2 + get_new_comment()
    if random.randint(1, 5) == 1:  # 1  of 5 times, insert a new line
        mod_2 = mod_2 + "\r\n"

mod_2 = mod_2 + "End Function"

# Module 3 - contains the CreateTextFile

mod_3_name = get_new_variable_name()
mod_3 = ""

mod_3 = mod_3 + "Function {}({} As {}) As {}\r\n".format(
    get_new_variable_name(),
    get_new_variable_name(),
    get_var_type(),
    get_var_type()
)

for i in range(1, random.randint(10, 12)):
    mod_3 = mod_3 + get_new_unused_variable_and_definition()
    if random.randint(1, 7) < 7:  # 6 of 7 times, insert a comment after
        mod_3 = mod_3 + get_new_comment()
    if random.randint(1, 5) == 1:  # 1  of 5 times, insert a new line
        mod_3 = mod_3 + "\r\n"

mod_3 = mod_3 + "End Function\r\n"

# real code time
# code from sample:
'''
Public Sub adWspA8(aSVp8zOe As String, aH6MX As String)
Dim ajT74 As Object
Set ajT74 = CreateObject("Scripting.FileSystemObject")
Dim amxgqjco As Object
Set amxgqjco = ajT74.CreateTextFile(aSVp8zOe, True, True)
amxgqjco.Write aH6MX
amxgqjco.Close
'''

vba_variables['DropFileFunctionName'] = get_new_variable_name()
vba_variables['DropFileFunction_arg1'] = get_new_variable_name()
vba_variables['DropFileFunction_arg2'] = get_new_variable_name()
vba_variables['DropFileFunction_FileSystemObject'] = get_new_variable_name()
vba_variables['DropFileFunction_CreateTextFileObject'] = get_new_variable_name()

mod_3 = mod_3 + '''Public Sub {}({} As String, {} As String)
Dim {} As Object
Set {} = CreateObject("Scripting.FileSystemObject")
Dim {} As Object
Set {} = ajT74.CreateTextFile({}, True, True)
{}.Write {}
{}.Close\r\n'''.format(
    vba_variables['DropFileFunctionName'],
    vba_variables['DropFileFunction_arg1'],
    vba_variables['DropFileFunction_arg2'],
    vba_variables['DropFileFunction_FileSystemObject'],
    vba_variables['DropFileFunction_FileSystemObject'],
    vba_variables['DropFileFunction_CreateTextFileObject'],
    vba_variables['DropFileFunction_CreateTextFileObject'],
    vba_variables['DropFileFunction_arg1'],
    vba_variables['DropFileFunction_CreateTextFileObject'],
    vba_variables['DropFileFunction_arg2'],
    vba_variables['DropFileFunction_CreateTextFileObject']
)

for i in range(1, random.randint(10, 12)):
    mod_3 = mod_3 + get_new_unused_variable_and_definition()
    if random.randint(1, 7) < 7:  # 6 of 7 times, insert a comment after
        mod_3 = mod_3 + get_new_comment()
    if random.randint(1, 5) == 1:  # 1  of 5 times, insert a new line
        mod_3 = mod_3 + "\r\n"

mod_3 = mod_3 + "End Sub"

if __name__ == "__main__":
    print("\r\nModule 1\r\n")
    print("Name of module: " + mod_1_name + "\r\n")
    print("-----")
    print(mod_1)
    print("-----")
    print("\r\nModule 1\r\n")
    print("Name of module: " + mod_2_name + "\r\n")
    print("-----")
    print(mod_2)
    print("-----")
    print("\r\nModule 1\r\n")
    print("Name of module: " + mod_3_name + "\r\n")
    print("-----")
    print(mod_3)
    print("-----")
