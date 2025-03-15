### checker.py
문서 내 통일되지 않은 폰트를 찾아 확인함
특정 폰트 외 다른 폰트를 모두 붉은 색으로 처리함
폰트의 종류
```python
import win32com.client as win32 # 한글 실행을 위한 모듈
from tkinter import Tk # 파일 선택을 위한 모듈
from tkinter.filedialog import askopenfilename # 파일 선택창과 관련된 모듈
import os



def start_hwp(visible=False, open_file=None):
    '''
    한글 파일을 실행하는 함수
    visible 기본 값은 False, True로 하면 한글 창의 띄워져 보이도록 if 문을 이용하여 설정
    open_file 기본 값은 None,  선택한 파일 경로를 넣을 수 있도록 설정
    '''
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
 
    if open_file:
        hwp.Open(open_file)
    else:
        pass
    return hwp

def select_file():
    win = Tk()  # GUI 실행하고
    win.withdraw()
    hwpx = askopenfilename(title="한글 파일을 선택해주세요. by 우혁쌤",   #파일 선택 창 맨위에 보이는 문구
                             initialdir=os.getcwd(),                   # 기본적으로 현재 폴더를 먼저 띄우게 설정
                             filetypes=[("한/글파일", "*.hwp *.hwpx")]) # 선택하는 파일의 종류를 제한
    win.quit()  # GUI 종료
    return hwpx

def change_letter_color(face, r, g, b):
    '''
    글꼴을 하나 선택 후 
    해당 글꼴의 색을 모두 rgb에 맞춰서 바꿈
    한양신명조는 한글 전용폰트(HTF) 이므로 hwp.FontType("HTF")을 사용
    대부분 사실 hwp.FontType("TTF")
    '''
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc")
#     hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("HTF")
    hwp.HParameterSet.HFindReplace.FindCharShape.FaceNameHangul = face
#     hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("HTF")
    hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = hwp.RGBColor(r, g, b)
    hwp.HParameterSet.HFindReplace.ReplaceMode = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

# def change_letter_color_all(r, g, b):
#     '''
#     한글의 버그를 이용(찾는 글꼴이 없으면 모든 글꼴을 선택함)
#     모든 글자의 색을 모두 rgb에 맞춰서  
#     '''
#     hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
#     hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc")
#     hwp.HParameterSet.HFindReplace.FindCharShape.FontTypeHangul = hwp.FontType("TTF")
#     hwp.HParameterSet.HFindReplace.FindCharShape.FaceNameHangul = "모두"
#     hwp.HParameterSet.HFindReplace.ReplaceCharShape.FontTypeHangul = hwp.FontType("TTF")
#     hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = hwp.RGBColor(r, g, b)
#     hwp.HParameterSet.HFindReplace.ReplaceMode = 1
#     hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
#     hwp.HParameterSet.HFindReplace.FindType = 1
#     hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

def check_letter(face, r, g, b):
    hwp.SetPos(0, 0, 0)
    hwp.Run("MoveSelDocBegin")
    hwp.InitScan()
    hwp.Run("MoveSelDocEnd")
    hwp.Run("CharShapeTextColorRed")
    change_letter_color(face, r, g, b)
    area = 2
    while True:
        hwp.SetPos(area, 0, 0)
        if hwp.GetPos()[0] == 0:
            break
        while True:
            hwp.Run("MoveSelDocBegin")
            hwp.InitScan()
            hwp.Run("MoveSelDocEnd")
            hwp.Run("CharShapeTextColorRed")
            change_letter_color(face, 0, 0, 0)
            area += 1
            hwp.SetPos(area, 0, 0)
            if hwp.GetPos()[0] == 0:
                break

if __name__ == '__main__':

    hwp = start_hwp(visible=True, open_file=select_file())
       
    # "한양신명조" 서체만 검게 나머지는 빨간색
    check_letter("한양신명조", 0, 0, 0)
```
### space_letter.py
글 작성중 단어가 줄을 넘어가는 경우, 글자의 자간을 자동으로 조정함.
### ascending.py
논문 등의 참고 문헌의 순서가 섞여있는 경우 오름차순으로 정렬
한글, 영어 순서로 정렬하도록 제작
중간에 separate_list 함수는 chatgpt를 이용하여 제작
```python

import win32com.client as win32
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os

def start_hwp(visible=False, open_file=None):
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
    if open_file:
        hwp.Open(open_file)
    else:
        pass
    return hwp

def select_file():
    win = Tk()  # GUI 실행하고
    win.withdraw()
    hwpx = askopenfilename(title="한글 파일을 선택해주세요. by 우혁쌤",
                             initialdir=os.getcwd(),
                             filetypes=[("한/글파일", "*.hwp *.hwpx")])
    win.quit()  # GUI 종료
    return hwpx

def InsertText(text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def separate_list(lst):
    eng_lst = []
    kor_lst = []
    for item in lst:
        if isinstance(item, str) and item[0].isalpha():
            if ord(item[0]) < 256: # 영어로 시작하는 경우
                eng_lst.append(item)
            else: # 한글로 시작하는 경우
                kor_lst.append(item)
    return eng_lst, kor_lst

if __name__ == '__main__':
    hwp = start_hwp(visible=True, open_file=select_file())

    content = []
    hwp.InitScan()
    while True:
        state, text = hwp.GetText()
        content.append(text)
        if state <= 1:
            break
    hwp.ReleaseScan()
 
    # 전처리(중복제거)
    content=set(content)
    content=list(content)
    # ''제거
    content.remove('')

    eng_lst, kor_lst=separate_list(content)

    # 정렬
    kor_lst.sort()
    eng_lst.sort()


    hwp.HAction.Run("MoveDocEnd")
    hwp.Run("BreakPage")
    InsertText("정렬된 문헌\r\n")

    for i in range(0, len(kor_lst)):
        text=kor_lst[i]
        InsertText(text)

    for i in range(0, len(eng_lst)):
        text=eng_lst[i]
        InsertText(text)

    hwp.SaveAs(Path=hwp.Path.replace(".hwp", "(정렬).hwp"), Format=hwp.XHwpDocuments.Item(0).Format)

```
### word_ascending.py
### ascending.py의 MS Word버전
### table_extractor.py
같은 형식을 가진 여러 문서에서 표의 위치에 대한 내용을 추출하여 모아주는 함수
현재 함수의 포멧은 하나의 표로 구성된 문서, 파일명(이름).hwp 파일에 대해 (이름)과 표의 특정 위치에 있는 내용을 뽑아서 통합.hwp에 저장함.
### pyinstaller.md
pyinstaller를 이용하여 파이썬 설치 없이 사용할 수 있도록 배포하는 법을 설명(예정)
### hwp_to_python.md
한글의 매크로 기록 기능을 이용하여 파이썬 코드로 변환하는 법(예정)

## 자간, 폰트, 폰트크기 자동화
