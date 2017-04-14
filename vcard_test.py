import vobject

j = vobject.vCard()

# j.add('BEGIN')
# j.BEGIN.value('vcard')
j.add('n')
j.n.value = vobject.vcard.Name(family='Ekaterina', given='Petergof')
j.add('fn')
j.fn.value = 'Ekaterina Petergof'
j.add('email')
j.email.value = 'Merkusheva@mkelite.ru'
j.email.type_param = "INTERNET"

# j.add('END')
# j.END.value('vcard')

# j.serialize()
j.prettyPrint()
# print(j)

with open('test', 'a') as f:
	f.write(str(j))