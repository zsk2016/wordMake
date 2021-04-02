import base64
import sys

def tu_zi_jie():
    with open('2.jpg','rb') as fp:
        tu = base64.b64encode(fp.read())
        zi_tu(tu)

def zi_tu(str):
    # b_tu = b'iVBORw0KGgoAAAANS....UhEU'
    tu_b = base64.b64decode(str)
    with open('tu.jpg', 'wb') as fp:
        fp.write(tu_b)

if __name__ == '__main__':
    str1 = sys.argv[1]
    tu_b = base64.b64decode(str1)
    with open("test.txt","wb") as f:
            f.write(tu_b)
    f.close()
    zi_tu(str1)
    # tu_zi_jie()