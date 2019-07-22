import string
import random
import os

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
    this_comment_word_count = random.randint(1, 5)
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


if __name__ == "__main__":
    example_output = ''
    for i in (range(1, 30)):
        example_output = example_output + get_new_unused_variable_and_definition()
        example_output = example_output + get_new_comment()
    print(example_output)