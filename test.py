f = open('signature.html', 'r',encoding='utf8')
a = f.read()
b= "%saaa"
print(a.format('a','3月'))