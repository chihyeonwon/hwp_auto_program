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


def num_letter():
    """
    자간자동조정 함수에서
    라인 끝에 걸쳐진 단어의
    앞뒤길이를 각각 계산하기 위함.
    """
    hwp.InitScan(Range=0xff)
    _, text = hwp.GetText()
    hwp.ReleaseScan()
    return len(text)


def adjustment():
    """
    모든 라인을 순회하면서
    끝에 걸쳐친 단어를 탐색함.

    잘린 단어의 앞이 길면
    라인 전체의 자간을 줄이고,

    잘린 단어의 뒤가 길면
    라인 전체의 자간을 늘임.

    한 줄 문단이 되거나
    걸쳐진 단어가 없으면 종료.
    """
    while True:
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveSelWordBegin")
        front_length = num_letter()
        if front_length == 0:  # 단어가 잘려있지 않으면 자간조정 중지
            break
        hwp.Run("MoveSelWordEnd")
        back_length = num_letter()
        if not (front_length and back_length):  # 한 줄 문단이면 자간조정 중지
            hwp.Run("Cancel")
            break
        hwp.Run("MoveWordBegin")

        # 자간 조정된 마지막 단어들을 표기하는 기능 추가
        # 필요에 따라서 주석처리하여 사용
        # 기능 1은 기울임과 밑줄 추가 (현재 주석 처리)
        # 기능 2는 빨간 형광펜 처리(필요에 따라 rgb 컬러 변경 가능)
        hwp.Run("MoveSelWordEnd")
#         charshape.Italic = True  # 이탤릭
#         charshape.UnderlineType = 1  # 밑줄
        hwp.HAction.GetDefault("MarkPenShape", hwp.HParameterSet.HMarkpenShape.HSet)
        hwp.HParameterSet.HMarkpenShape.Color = hwp.RGBColor(255, 0, 0)
        hwp.HAction.Execute("MarkPenShape", hwp.HParameterSet.HMarkpenShape.HSet)
        hwp.HAction.Run("Cancel")
        hwp.Run("MoveWordBegin")
        
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveSelLineBegin")
        if front_length >= back_length:  # 앞이 길면
            hwp.Run("CharShapeSpacingDecrease")  # 라인 자간 -1%
        else:  # 뒤가 길면
            hwp.Run("CharShapeSpacingIncrease")  # 라인 자간 +1%
        hwp.Run("Cancel")


def ctrl_adjustment():
    """
    표나 글상자 등 텍스트가 들어가는
    모든 영역의 자간을 조정하기 위함
    """
    area = 2
    while True:
        hwp.SetPos(area, 0, 0)
        if hwp.GetPos()[0] == 0:
            break
        while True:
            adjustment()
            hwp.Run("MoveLineEnd")
            hwp.Run("MoveNextPosEx")
            if hwp.GetPos()[0] == 0:
                break
        area += 1


def end_position():
    """
    본문 탐색 while문의 종료 조건으로
    "문서 끝에 도착하면 반복종료"를 구현하기 위해
    문서 끝 위치를 미리 추출해 둠
    """
    hwp.Run("MoveDocEnd")
    end_pos = hwp.GetPos()  # 종료위치 저장
    hwp.Run("MoveDocBegin")
    return end_pos

def select_file():
    win = Tk()  # GUI 실행하고
    win.withdraw()
    hwpx = askopenfilename(title="한글 파일을 선택해주세요. by 우혁쌤",
                             initialdir=os.getcwd(),
                             filetypes=[("한/글파일", "*.hwp *.hwpx")])
    win.quit()  # GUI 종료
    return hwpx

if __name__ == '__main__':
    hwp = start_hwp(visible=True, open_file=select_file())
    end = end_position()

    # 본문 자간조정
    while hwp.GetPos() != end:
        adjustment()
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveNextPosEx")

    # 표 및 글상자 자간조정
    ctrl_adjustment()
    # print("자간조정 작업 끝!")
    
    hwp.SaveAs(Path=hwp.Path.replace(".hwp", "(자간조정).hwp"), Format=hwp.XHwpDocuments.Item(0).Format)
```
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
```python
import os
from tkinter.filedialog import askopenfilenames
import win32com.client as win32

def start_hwp(visible=False, open_file=None):
    '''
    한글 파일을 시작하도록 하는 함수
    visible=False # 한글 창을 보이지 않도록 설정 (True로 변경시 한글 창이 보임)
    '''
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
    if open_file:
        hwp.Open(open_file)
    else:
        pass
    return hwp

def choose_file():
    """
    파일선택 함수
    """
    filelist = askopenfilenames(title="수정할 한/글문서를 모두 선택해주세요.",
                     initialdir=os.getcwd(),
                     filetypes=[("한/글파일", "*.hwp *.hwpx")])
    return filelist

def get_name(path):
    '''
    파일 이름이 --(이름) 꼴일때
    (이름)을 뽑아내는 함수
    '''
    start = path.find("(")
    end = path.find(")") + 1
    result = path[start:end]
    return result

def get_text(hwp):
    '''
    한글에서 적혀있는 글자를 추출하는 함수
    '''
    hwp.InitScan(Range=0xff)
    total_text = ""
    state = 2
    while state not in [0, 1]:
        state, text = hwp.GetText()
        total_text += text
    hwp.ReleaseScan()
    return total_text

def InsertText(text):
    '''
    한글에서 글을 입력하는 함수
    '''
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def move_to_begin(hwp):
    hwp.HAction.Run("MoveDocBegin")  # 문서 시작으로 이동

if __name__ == '__main__':              # 실행
    contents = []                       # 내용을 담을 리스트
    names=[]                            # 이름을 담을 리스트
    filelist=choose_file()              # 파일선택
    for file in filelist:               # 선택한 파일들 중에 하나씩 선택해서
        hwp = start_hwp(open_file=file) # 파일을 열고
        move_to_begin(hwp)              # 시작점으로 돌아간 후
        path = hwp.Path                 # 경로를 추출하고(최종 저장할 파일의 경로를 같에 만들어주기 위해)
        name=get_name(hwp.Path)         # 해당 파일의 이름을 추출하고
        names.append(name)              # names 리스트에 담는다.
        ctrl = hwp.HeadCtrl             # 컨트롤을 활성화하고
        hwp.FindCtrl()                  # 컨트롤을 찾고
        hwp.Run("ShapeObjTableSelCell") # 첫 번째 셀로 진입
        for i in range(8):              # range(n) : n은 첫번째 셀에서 F5를 누르고 목표 셀까지 오른쪽 키로 이동한 횟수
            hwp.HAction.Run("TableRightCell") # 오른쪽 키보드를 누른 것과 같은 효과
        hwp.Run("ShapeObjTableSelCell") # 도착한 셀에 진입해서
        contents.append(get_text(hwp))  # 글자를 가져오고
        hwp.Clear(option=1)             # 파일을 닫음(새로 파일을 열기 위함)
    hwp.Quit()                          # 한글 일단 종료


    # 파일 이름만 추출
    directory, filename = os.path.split(path)

    hwp = start_hwp()                   # 한글 새로 시작
    for i in range(len(names)):         # names 리스트에 담긴 이름만큼 (제출한만큼)
        text_name=names[i]              # 이름을 선택하고
        text=contents[i]                # 내용을 선택하고
        InsertText(text_name)           # 이름을 적고
        InsertText("\r\n")              # 한줄 띄고
        InsertText(text)                # 내용을 적고
        InsertText("\r\n")              # 한줄 띄고
        InsertText("\r\n")              # 한줄 띄고
    
    hwp.SaveAs(Path=directory+'\\'+'통합.hwp', Format=hwp.XHwpDocuments.Item(0).Format)  # 저장하고 
    hwp.Clear() # 문서 닫기
    hwp.Quit() # 한글 종료
```
### pyinstaller.md
pyinstaller를 이용하여 파이썬 설치 없이 사용할 수 있도록 배포하는 법을 설명(예정)
### hwp_to_python.md
한글의 매크로 기록 기능을 이용하여 파이썬 코드로 변환하는 법(예정)

## 자간, 폰트, 폰트크기 자동화
