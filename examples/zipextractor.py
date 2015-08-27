__author__ = 'Shiva'

import zipfile

zfile = zipfile.ZipFile("evil.zip")

your_list = '12'
complete_list = []

for current in range(1,2):
    a = [i for i in your_list]
    for y in range(current):
        a = [x+i for i in your_list for x in a]
    complete_list = complete_list+a
    #if(complete_list[current] == 'test@123'):
    print(complete_list.__len__())

#complete_list = ['test123']
for item in complete_list.__iter__():


    item = item.encode('utf-8')
    try:
        zfile.extractall(pwd=item)
        print(item)
    except Exception as e:
        pass


#passFile = open('dictionary.txt')
#for line in passFile.readlines():
 #   password = line.strip('\n')
  #  try:
   #     zfile.extractall(pwd=password)
    #    print(password)

   # except Exception as e:
    #    print (e)