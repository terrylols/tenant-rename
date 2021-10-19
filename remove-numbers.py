import re
import os

print("Enter name of the file")
fileinp = input()
if not os.path.isfile(fileinp):
    print("No such file found")
    os.system('pause')
    exit()
print("Remove duplicates? 'y' for yes (default is no)")
choice = input()
print("Remove \"itw\" and \"it\"? 'y' for yes (default is no)")
choice2 = input()
print("Print for Excel or Word? 'excel' 'word' (There is no default here, you have to specify)")
choice1 = input()

    
yes = ['yes', 'y', 'Y', 'Yes', 'YES', 'YEs', 'yES', 'yeS', ]
no = ['n', 'no', 'NO', 'No', 'nO', 'N']

file1 = open(fileinp, "r")
a = file1.read().splitlines()
nums = "1234567890"
hui = []
for number_of_word in range(len(a)):
    word = a[number_of_word]
    if word[-1] in nums:
        asdf = [m.start() for m in re.finditer('-', word)]
        hui.append(word[:max(asdf)-2])
    else:
        hui.append(word)

if choice in yes:
    hui = list(dict.fromkeys(hui))

if choice2 in yes:
    for word in range(len(hui)):
        newword = hui[word]
        if not newword.find('itw') or not newword.find('it'):
            yeet = newword.find('-')
            hui[word] = newword[yeet+1:]
        else:
            continue

os.system('cls')

if choice1 == 'word':
    os.system('cls')
    print(', '.join(hui))
    os.system("pause")
elif choice1 == 'excel':
    os.system('cls')
    for line in range(len(hui)):
        print(hui[line])
    os.system("pause")

file1.close()
