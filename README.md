### checker.py
문서 내 통일되지 않은 폰트를 찾아 확인함
특정 폰트 외 다른 폰트를 모두 붉은 색으로 처리함
폰트의 종류
### space_letter.py
글 작성중 단어가 줄을 넘어가는 경우, 글자의 자간을 자동으로 조정함.
### ascending.py
논문 등의 참고 문헌의 순서가 섞여있는 경우 오름차순으로 정렬
한글, 영어 순서로 정렬하도록 제작
중간에 separate_list 함수는 chatgpt를 이용하여 제작
```

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
