import os

import PySimpleGUI as gui
import base64
import textwrap
from datetime import datetime, date
import numpy as np
from PIL import ImageFont, ImageDraw, Image
import pandas

import cv2

print(cv2.__file__)

BOLD = "bold"
ITALIC = "italic"
NORMAL = ""


def font(size, style):
    return "Arial", str(size), style


content = [
    [gui.Text("NEU SMART CHECKIN by SITDE", font=font(18, BOLD))],
    [gui.HSeparator()],
    [gui.Text("Thiết lập", font=font(16, NORMAL))],
    [gui.Column([
        [gui.Text(
            "1. Nhập danh sách sinh viên đăng ký tham gia.\nLưu ý danh sách này cần theo định dạng trong hướng dẫn.",
            font=font(12, NORMAL)),
            gui.Column(
                [
                    [gui.In(default_text="Chưa chọn", disabled_readonly_background_color="#c3c3c3",
                            size=(25, 1), enable_events=True, key='attendant_list', disabled=True),
                     gui.FileBrowse(button_text="Chọn tệp", key="select_attendant_list",
                                    file_types=[("Excel Files", ".xlsx")])]
                ]),
        ],
        [gui.Text(
            "2. Nhập danh sách thông tin các đối tượng tham gia chương trình.\nLưu ý danh sách này cần theo định dạng trong hướng dẫn.",
            font=font(12, NORMAL)),
            gui.Column(
                [
                    [gui.In(default_text="Chưa chọn", disabled_readonly_background_color="#c3c3c3",
                            size=(25, 1), enable_events=True, key='welcome_message', disabled=True),
                     gui.FileBrowse(button_text="Chọn tệp", key="select_welcome_message",
                                    file_types=[("Excel Files", ".xlsx")])]
                ]),
        ],
        [gui.Text("3. Chọn chế độ",
                  font=font(12, NORMAL)),
         gui.Combo(["Check-in", "Check-out"], default_value="Check-in", font=font(12, NORMAL), enable_events=True,
                   key="select_mode"),
         ],
        [gui.Text(
            "4. Chọn đường dẫn lưu file kết quả",
            font=font(12, NORMAL)),
            gui.Column(
                [
                    [gui.In(default_text="Chưa chọn", disabled_readonly_background_color="#c3c3c3",
                            size=(25, 1), enable_events=True, key='output_dist', disabled=True),
                     gui.FolderBrowse(button_text="Chọn đường dẫn")]
                ]),
        ],
        [gui.Text(
            "5. Cho phép người chưa có tên trong danh sách đăng kí vẫn có thể checkin/checkout",
            font=font(12, NORMAL)),
            gui.Checkbox(text="", key="allow_unregistered", enable_events=True),
        ],
    ], element_justification="left")],
    [gui.HSeparator()],
    [gui.Text("Hành động", font=font(16, NORMAL))],
    [gui.Button("BẮT ĐẦU CHẠY", key="start", font=font(16, BOLD))],
    [gui.Button("Tạo mã QR theo danh sách", button_color="#2d4373"), gui.Button("Thoát", button_color="#e85235")]
]

layout = [[gui.VPush()],
          [gui.Push(), gui.Column(content, element_justification='center'), gui.Push()],
          [gui.VPush()]]

window = gui.Window(title='NEU Smart Checkin by SITDE', layout=layout, size=(1224, 776))

attendantDF = None
welcomeMessages = None
outputDist = None
mode = "Check-in"
allow_unregistered = False

applicationStartDate = datetime.today().strftime("%d_%m_%Y_%X").replace(":", "")
mode = mode.replace("-", "").lower()

df = None
CHECK_MSV = []
CHECK_TIME = []
CHECK_TARGET = []

while True:
    event, values = window.read()
    print(event, values)

    if event == "attendant_list":
        path = values["attendant_list"]
        attendantDF = pandas.read_excel(path)
        gui.PopupOK("Chọn file chứa danh sách người tham gia thành công. Đã phát hiện "
                    + str(len(attendantDF.index)) + " người đăng ký tham gia", title="THÀNH CÔNG")
    elif event == "welcome_message":
        path = values["welcome_message"]
        welcomeMessages = pandas.read_excel(path)
        gui.PopupOK("Chọn file danh sách đối tượng tham gia thành công. Đã phát hiện "
                    + str(len(welcomeMessages.index)) + " đối tượng", title="THÀNH CÔNG")
    elif event == "output_dist":
        path = values["output_dist"]
        outputDist = path
        gui.PopupOK(
            "Chọn đường dẫn chứa file kết quả checkin/checkout thành công. File sẽ được lưu dưới dạng <ngày_hôm_nay>_<checkin/checkout>.xlsx",
            title="THÀNH CÔNG")
    elif event == "select_mode":
        mode = values["select_mode"]
        mode = mode.replace("-", "").lower()
        gui.PopupOK("Đã đổi chế độ sang " + mode, title="THÀNH CÔNG")
    elif event == "allow_unregistered":
        allow_unregistered = values["allow_unregistered"]

    if event == 'Thoát' or event == gui.WINDOW_CLOSED:
        break

    if event == "start":
        break
        # window.close()

studentDataXLSX = attendantDF
studentData = dict(zip(studentDataXLSX["Mã sinh viên"].map(lambda e: str(e)),
                       zip(studentDataXLSX["Tên"], studentDataXLSX["Đối tượng"])))

messageData = welcomeMessages
messageMap = dict(zip(messageData["Đối tượng"], zip(messageData["Lời chào checkin"], messageData["Lời chào checkout"],
                                                    messageData["Checkin tối đa"],
                                                    messageData["Checkout tối đa"])))

def findStudent(studentId):
    if studentId in studentData:
        _data = list(studentData[studentId])
        _target = _data[1]
        _data.append(messageMap[_target][2])
        _data.append(messageMap[_target][3])
        return _data
    return None


def findWelcomeMessage(target, mode):
    if target in messageMap:
        if mode == "checkin":
            return messageMap[target][0]
        else:
            return messageMap[target][1]
    return None


def getMessage(mapvalue: dict, target):
    message = findWelcomeMessage(target, mode)
    if message is None:
        return "INVALID TARGET"
    return message.replace("{{MSV}}", mapvalue["MSV"]).replace("{{Name}}", mapvalue["Name"])


def appendData(key, value, target):
    CHECK_MSV.append(key)
    CHECK_TIME.append(value)
    CHECK_TARGET.append(target)

    df = pandas.DataFrame()
    df["Mã sinh viên"] = CHECK_MSV
    df["Thời gian"] = CHECK_TIME
    df["Đối tượng"] = CHECK_TARGET

    df.to_excel(outputDist + "/" + applicationStartDate + "_" + mode + ".xlsx")
    return True


capture = cv2.VideoCapture(0)
qrDetector = cv2.QRCodeDetector()

window_name = "NEU Smart Checkin"
cv2.namedWindow(window_name, cv2.WND_PROP_FULLSCREEN)
cv2.setWindowProperty(window_name, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

font = ImageFont.truetype("fonts/RobotoFlex.ttf", 20)

content = "Đang chờ " + mode + " ..."
lastStudentID = ""
lastCheckinTimestamp = None
delaying = False

checkCount = {}

while True:
    try:
        _, img = capture.read()
        data, __, _ = qrDetector.detectAndDecode(img)

        if data and not delaying:
            lastCheckinTimestamp = datetime.now()

            rawData = ""
            rawDataSplitted = []
            try:
                rawData = base64.b64decode(data.encode("ascii")).decode("ascii")
                rawDataSplitted = rawData.split("_")
            except:
                lastStudentID = ""
                content = "Mã QR của bạn không hợp lệ"

            if len(rawDataSplitted) != 2:
                lastStudentID = ""
                content = "Mã QR của bạn không hợp lệ"
            else:
                lastStudentID = rawDataSplitted[1]

                _student = findStudent(lastStudentID)

                unregisterState = True
                unregisterSuffix = ""
                if _student is None:
                    if allow_unregistered:
                        _student = ["", list(messageMap.keys())[0], 1, 1]
                        unregisterSuffix = "_KĐK"
                    else:
                        content = "Bạn chưa đăng ký tham gia chương trình này, do đó, bạn không được checkin."
                        unregisterState = False

                if unregisterState:
                    able = True
                    if rawData in checkCount:
                        if (mode == "checkin" and checkCount[rawData] >= int(_student[2])) \
                                or (mode == "checkout" and checkCount[rawData] >= int(_student[3])):
                            able = False

                    if not able:
                        content = "Mã sinh viên " + lastStudentID + " đã " + mode + " thành công trước đó. Không cần " + mode + " lại."
                    else:
                        if rawData in checkCount:
                            checkCount[rawData] += 1
                        else:
                            checkCount[rawData] = 1

                        _m = getMessage(mapvalue={"MSV": lastStudentID, "Name": _student[0]},
                                        target=rawDataSplitted[0])
                        if _m == "INVALID TARGET":
                            lastStudentID = ""
                            content = "Bạn không phải đối tượng có thể tham gia checkin/checkout chương trình này."
                        else:
                            content = _m
                            appendData(lastStudentID, lastCheckinTimestamp.strftime("%X"), rawDataSplitted[0] + unregisterSuffix)

                    delaying = True
        else:
            if lastCheckinTimestamp is not None:
                delay = datetime.combine(date.today(), datetime.now().time()) - datetime.combine(date.today(),
                                                                                                 lastCheckinTimestamp.time())
                if delay.seconds > 3:
                    lastCheckinTimestamp = None
                    delaying = False
                    lastStudentID = ""
                    content = "Đang chờ " + mode + " ..."

        notification = np.zeros((200, 400, 3), np.uint8)
        b, g, r, a = 0, 255, 0, 0

        img_pil = Image.fromarray(notification)
        draw = ImageDraw.Draw(img_pil)

        offset = 0
        wrapText = textwrap.wrap(content, width=40)
        wrapText.reverse()
        wrapText.append(lastStudentID)
        for line in wrapText:
            w = draw.textlength(line, font=font)
            W, H = img_pil.size
            x, y = 0.5 * (W - w), 0.90 * H - 20
            draw.text((x, y - offset), line + "\n", font=font, fill=(b, g, r, a))
            offset += 30

        notification = np.array(img_pil)

        cv2.imshow("Thong bao " + mode, notification)

        cv2.imshow(window_name, img)
        if cv2.waitKey(1) == ord('q'):
            break
    except:
        print("[WARNING]")

capture.release()
cv2.destroyAllWindows()
