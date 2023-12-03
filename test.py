txt = "Částka: 200 500 CZK"
erasers = ["Částka:","CZK"," "]
for eraser in erasers:
    txt = txt.replace(eraser,"")
print(txt)