f = open(r'C:\Users\LJY\Desktop\forcool-ad\src\App.jsx', encoding='utf-8')
c = f.read()
f.close()
v = [l for l in c.split('\n') if 'v3.' in l or 'v2.' in l]
print(v[:2])
print('lines:', c.count('\n'))
