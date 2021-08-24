import json
import re
from time import sleep
from os import system


def clear():
    system('cls')


class EditDictMenu:
    key_list = []

    def __init__(self):
        try:
            with open('options/key_words_dict.json', 'r') as json_file:
                self.key_list = json.load(json_file)
        except FileNotFoundError:
            self.key_list = []
        except Exception as e:
            print(e.args)

    def save(self):
        with open('options/key_words_dict.json', 'w') as json_file:
            json.dump(self.key_list, json_file)

    def show_dict(self):
        for i in range(0, len(self.key_list)):
            dct = self.key_list[i]
            print('{}) Key words: {}'.format(i+1, dct['key_words']))
            print('   Replaced values:')
            for key in dct['replace_keys'].keys():
                print('     {}: {}'.format(dct['replace_keys'][key], key))
            print('   choose by: {}'.format(dct['sort_values']))

    def add_dict(self):
        print('Adding new key dict:')
        result_dict = {}
        key_words = input('Input ALL table fields from left to right separating with semicolon: ')
        wrds = []
        for word in re.split(';', key_words):
            wrds.append(word)
        result_dict.update({'key_words': wrds})
        result_dict.update({'replace_keys': {}})
        print('Input replaced fields:')
        result_dict['replace_keys'].update({input(' sort_val: '): 'sort_val'})
        result_dict['replace_keys'].update({input(' instance: '): 'instance'})
        result_dict['replace_keys'].update({input(' value: '): 'value'})
        result_dict['replace_keys'].update({input(' variable_name: '): 'variable_name'})
        key_words = input('Input chosen values in sort_val thought semicolon: ')
        result_dict.update({'sort_values': re.split(';', key_words)})

        self.key_list.append(result_dict)

    def delete_dict(self, num):
        if len(self.key_list) >= num >= 1:
            self.key_list.pop(num-1)
        else:
            print('No such num dict !')
            sleep(2)

    def run(self):
        loop = True
        while loop:
            self.show_dict()
            command = input('Commands:\n add: add a new key dict\n delete <num>: delete key dict\n s: save\n q: quit\n')
            if command == 'add':
                self.add_dict()
            elif re.match('delete', command):
                try:
                    self.delete_dict(int(re.split(' ', command)[1]))
                except Exception as e:
                    print(e)
            elif command == 's':
                self.save()
            elif command == 'q':
                self.save()
                loop = False
            clear()


edit = EditDictMenu()
edit.run()

