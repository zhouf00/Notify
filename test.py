f = open('signature.html', 'r', encoding='utf8')
a = f.read()
print(a.format('称谓','3月','名字','email', 'email', '手机'))