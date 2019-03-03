from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5 import QtCore, QtGui
from subjectInformation import Ui_subinfo
from MIST import Ui_Form
from psychopy import visual, core, event, clock
import random
import xlrd
import xlsxwriter
import time
import numpy
import sys
import os
from win32con import WM_INPUTLANGCHANGEREQUEST
import win32gui
import win32api

global MeanT
global MIST_time
global correctRate
global dotexp_picnum
global block_num
MeanT = 10  # 单位为s，练习阶段结束后会获取被试平均时间，初始值不会影响结果，不为0是为了调试
MIST_time = 3   # 单位为minutes，实验前请自行修改
correctRate = 0.6  # 练习阶段点探测任务正确率限制
dotexp_picnum = 48  # 正式试验点探测任务呈现图片对数，勿随意修改
block_num = 3  # block的数量，小于3时没问题，大于3时在点侦测任务的呈现时会出现问题
fixation_duration = [0.5, 1]  # 注视点呈现时间范围为0.5-1s


class subInfoWindow(QWidget, Ui_subinfo):  # 呈现被试信息框的类

    acceptIsClicked = False

    def __init__(self):
        super(subInfoWindow, self).__init__()
        self.setupUi(self)
        self.initial()

    def initial(self):
        self.lineEdit_2.setValidator(QtGui.QIntValidator())  # 设置年龄只能输入整数

    def accept(self):
        """点击ok键对应的事件
        """
        ID = self.lineEdit.text()
        group = 0
        order = 0
        gender = 0
        if self.checkBox.isChecked() and not self.checkBox_2.isChecked():
            group = 1
        if self.checkBox_2.isChecked() and not self.checkBox.isChecked():
            group = 2

        if self.checkBox_3.isChecked() and not self.checkBox_4.isChecked():
            order = 2
        if self.checkBox_4.isChecked() and not self.checkBox_3.isChecked():
            order = 1

        if self.checkBox_5.isChecked() and not self.checkBox_6.isChecked():
            gender = 2
        if self.checkBox_6.isChecked() and not self.checkBox_5.isChecked():
            gender = 1

        age = self.lineEdit_2.text()
        global subinfo
        subinfo = [ID, group, order, gender, age]
        print(subinfo)

        print('被试信息填写成功，开始实验!')

        self.close()

    def reject(self):
        """点击cancel键对应的事件"""
        print('被试信息填写不成功，取消实验！')
        sys.exit()


class problem_presentThread(QtCore.QThread):  # MIST任务中用于检测超时的子线程

    update_signal = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super(problem_presentThread, self).__init__(parent)
        self.stop_flag = 0

    def run(self):
        time_stamp = time.time()
        while self.stop_flag == 0 and time.time()-time_stamp < MeanT:
            pass
        if self.stop_flag == 0:
            self.update_signal.emit('overtime')  # 如果在MeanT之内未回答，则发出overtime信号，否则发出intime信号
        else:
            self.update_signal.emit('intime')


class MISTWindow(QWidget, Ui_Form):  # 用于MIST任务呈现的类

    problem_id = 0
    problems = []
    mean_time = 1.5
    time_change_flag = 0
    continue_right_count = 0
    continue_wrong_count = 0
    block = 0
    group = 0
    work_thread = None  # 用于判断是否超时的子线程
    correct_num = 0

    def __init__(self, block, group):
        super(MISTWindow, self).__init__()
        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.FramelessWindowHint)
        self.showFullScreen()
        self.progressBar.setStyleSheet("""QProgressBar {border: 2px solid grey;border-radius: 5px;} 
            QProgressBar::chunk {background-color: #C0C0C0;width: 20px;}""")
        self.pushButton.clicked.connect(self.number_input)
        self.pushButton_2.clicked.connect(self.number_input)
        self.pushButton_3.clicked.connect(self.number_input)
        self.pushButton_4.clicked.connect(self.number_input)
        self.pushButton_5.clicked.connect(self.number_input)
        self.pushButton_6.clicked.connect(self.number_input)
        self.pushButton_7.clicked.connect(self.number_input)
        self.pushButton_8.clicked.connect(self.number_input)
        self.pushButton_9.clicked.connect(self.number_input)
        self.pushButton_10.clicked.connect(self.number_input)
        self.pushButton_11.clicked.connect(self.close)
        self.pushButton_11.setVisible(False)
        self.block = block
        self.group = group
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.progressbar_timer)
        self.step = 0
        self.timer.start(MIST_time*60)  # 将MIST_time分成1000份并转化为ms，使得进度条值与时间对应
        self.problems = self.math_problem_read()
        if block in [1, 2, 3]:
            self.problem_id = 50 * block  # 使block1从第51题开始，block2从101题开始...
            self.mean_time = MeanT

        self.problem_present(self.problem_id)

    def problem_present(self, id):
        """
        呈现算术题
        """
        problem_str = ' '.join(self.problems[id]+'=?')
        self.lineEdit.setText(problem_str)
        # self.lineEdit.setStyleSheet("font-size: 100px")
        self.lineEdit.setAlignment(QtCore.Qt.AlignCenter)
        if self.block in [1, 2, 3] and self.group != 2:  # 如果是正式实验且为压力组，启动另一个进程
            self.work_thread = problem_presentThread()
            self.work_thread.start()
            self.work_thread.update_signal.connect(self.problem_presentation_thread)

    def problem_presentation_thread(self, str):
        """
        work_thread的信号接收函数
        """
        if str == 'overtime':  # 收到overtime信号时，出现超时提醒并呈现新题
            self.lineEdit_2.setText('超时!')
            self.problem_id = self.problem_id + 1
            self.problem_present(self.problem_id)
            # print('平均答题时间为{}'.format(MeanT))

    def progressbar_timer(self):
        """
        定时改变进度条值，同时如果时间用完，跳转到结束状态 
        """
        if self.block in [1, 2, 3] and self.group != 2:
            if self.continue_right_count >= 3:  # 如果连续作对三题，就把进度条值-时间增加10%
                self.step = self.step + 100
                self.continue_right_count = 0
            if self.continue_wrong_count >= 3:  # 如果连续做错三题，就把进度条值-时间减少10%
                self.step = self.step - 100
                self.continue_wrong_count = 0

        self.step = self.step + 1
        self.progressBar.setValue(self.step)

        if self.step >= 1000:  # 进度条值最大被设为1000，这意味着时间用光了

            if self.block in [1, 2, 3] and self.group != 2:  # 在正式实验且为压力组中，如果存在子线程，将其终止
                self.work_thread.stop_flag = 1

            self.timer.stop()

            if self.correct_num > 0 and self.block == 0:  # 如果被试至少做对了1道题，并且处于练习阶段，计算平均时间
                self.mean_time = MIST_time*60 / self.correct_num

            self.lineEdit.setText(u'按continue键继续')
            self.lineEdit_2.setText('TIME OUT!')
            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(False)
            self.pushButton_3.setEnabled(False)
            self.pushButton_4.setEnabled(False)
            self.pushButton_5.setEnabled(False)
            self.pushButton_6.setEnabled(False)
            self.pushButton_7.setEnabled(False)
            self.pushButton_8.setEnabled(False)
            self.pushButton_9.setEnabled(False)
            self.pushButton_10.setEnabled(False)
            self.pushButton_11.setEnabled(True)
            self.pushButton_11.setVisible(True)

    def math_problem_read(self):
        """用于读取已生成的数学题
        """      
        problem_book = xlrd.open_workbook('math_problems.xlsx')
        sheet1 = problem_book.sheet_by_name('sheet1')
        problems = sheet1.col_values(0) + sheet1.col_values(1)  # 读取第1,2列算术题
        random.shuffle(problems)
        return problems

    def number_input(self):
        """
        与界面数字键盘关联的函数，获取被点击按钮的信息
        """
        button = self.sender()  # 获取发送信号部件的信息，即界面上的数字键盘
        # print(button.objectName())
        number_input = int(button.text())
        solution = eval(self.problems[self.problem_id])
        if self.group == 1 or self.block == 0:  # 压力组进行反馈
            if number_input == solution:
                self.lineEdit_2.setText(u'正确!')
                self.correct_num = self.correct_num + 1
                self.continue_right_count = self.continue_right_count + 1
                self.continue_wrong_count = 0
            else:
                self.lineEdit_2.setText(u'不正确: 答案为{}'.format(solution))
                self.continue_right_count = 0
                self.continue_wrong_count = self.continue_wrong_count + 1

            if self.block in [1, 2, 3]:  # 在正式实验中，按键即表明非超时，停止子线程
                self.work_thread.stop_flag = 1

        self.problem_id = self.problem_id + 1  # 呈现下一题
        self.problem_present(self.problem_id)


def main_window_state(win, state):
    """
    用于改变psychopy窗口的优先级，保证Qt窗口能出现在最上层
    """
    if state == 1:
        win.winHandle.maximize()
        win.winHandle.set_fullscreen(True)
        win.winHandle.activate()
        win.flip()
    elif state == 2:
        win.winHandle.minimize()  # minimise the PsychoPy window
        win.fullscr = False  # disable fullscreen
        win.flip()  # redraw the (minimised) window


def intro_img_show(win, imgPath):
    """
    此函数用于呈现指示语图片
    """
    img = visual.ImageStim(win, image=imgPath, size=[1.5, 1.5])
    img.draw()
    win.flip()
    key = event.waitKeys(keyList=['space', 'escape'])
    if key[0] == 'space':
        pass
    else:
        print("按下esc键，程序退出!")
        sys.exit()


def pra_dotexp(win):
    """
    此函数用于练习阶段的点侦测任务
    """
    correct_times = 0
    pic_path = '.\\Pic\\Pra_Pic'
    files = os.listdir(pic_path)  # 读取该文件夹下所有文件的文件名

    for i in range(0, 10):

        fixation = visual.TextStim(win, text='+', pos=[0, 0], height=0.3, color=(-1, -1, -1))  # 呈现注视点
        fixation.draw()
        win.flip()
        clk = clock.CountdownTimer(random.uniform(fixation_duration[0], fixation_duration[1]))  # 注视点随机呈现500-1000ms
        while clk.getTime() > 0:
            continue

        pic_left_path = pic_path + '\\' + files[2*i]
        pic_right_path = pic_path + '\\' + files[2*i+1]
        pic_left = visual.ImageStim(win, image=pic_left_path, size=[1, 1], pos=[-0.5, 0])  # 呈现左右两张图片
        pic_right = visual.ImageStim(win, image=pic_right_path, size=[1, 1], pos=[0.5, 0])
        pic_left.draw()
        pic_right.draw()
        fixation.draw()
        win.flip()

        clk = clock.CountdownTimer(1)
        while clk.getTime() > 0:
            continue

        dotposition = random.choice([-0.5, 0.5])  # 点随机呈现在左右侧
        dot = visual.TextStim(win, text='·', pos=[dotposition, 0], height=0.5, color=(-1, -1, -1))
        dot.draw()
        fixation.draw()
        win.flip()

        keydown = event.waitKeys(keyList=['f', 'j'])

        # 根据按键反应进行反馈
        if (keydown[0] == 'f' and dotposition == -0.5) or (keydown[0] == 'j' and dotposition == 0.5):
            feedback = visual.TextStim(win, text='正确', pos=[0, 0], height=0.3, color=(-1, -1, -1))
            feedback.draw()
            win.flip()
            correct_times = correct_times + 1
        else:
            feedback = visual.TextStim(win, text='错误', pos=[0, 0], height=0.3, color=(-1, -1, -1))
            feedback.draw()
            win.flip()

        core.wait(1)

    return correct_times/10  # 返回正确率


def dotexp(win, block):
    """
    此函数用于正式实验阶段的点侦测任务
    """
    dotexp_result = [0] * dotexp_picnum
    reaction_time = [0] * dotexp_picnum
    pic_pair_class = [0] * dotexp_picnum  # 储存每对呈现图片的种类，0=LN, 1=HN, 2=NN, 3=HL
    files1 = []
    files2 = []

    intro_img_show(win, imgPath='.\\pic\\intro\\Exp_Dot_Intro.png')
    # 根据block读取图片文件名称
    if block == 1:
        LN_pic_path_L = '.\\Pic\\Exp_Pic\\Block1\\Health-Nofood\\Health'
        HN_pic_path_H = '.\\Pic\\Exp_Pic\\Block1\\Unhealth-Nofood\\Unhealth'
        NN_pic_path_N1 = '.\\Pic\\Exp_Pic\\Block1\\Nofood-Nofood\\Nofood1'
        HL_pic_path_L = '.\\Pic\\Exp_Pic\\Block1\\Health-Unhealth\\Health'
        files1 = os.listdir(LN_pic_path_L) + os.listdir(HN_pic_path_H) + os.listdir(NN_pic_path_N1) + os.listdir(
            HL_pic_path_L)
        LN_pic_path_N = '.\\Pic\\Exp_Pic\\Block1\\Health-Nofood\\Nofood'
        HN_pic_path_N = '.\\Pic\\Exp_Pic\\Block1\\Unhealth-Nofood\\Nofood'
        NN_pic_path_N2 = '.\\Pic\\Exp_Pic\\Block1\\Nofood-Nofood\\Nofood2'
        HL_pic_path_H = '.\\Pic\\Exp_Pic\\Block1\\Health-Unhealth\\Unhealth'
        files2 = os.listdir(LN_pic_path_N) + os.listdir(HN_pic_path_N) + os.listdir(NN_pic_path_N2) + os.listdir(
            HL_pic_path_H)
    elif block == 2:
        LN_pic_path_L = '.\\Pic\\Exp_Pic\\Block2\\Health-Nofood\\Health'
        HN_pic_path_H = '.\\Pic\\Exp_Pic\\Block2\\Unhealth-Nofood\\Unhealth'
        NN_pic_path_N1 = '.\\Pic\\Exp_Pic\\Block2\\Nofood-Nofood\\Nofood1'
        HL_pic_path_L = '.\\Pic\\Exp_Pic\\Block2\\Health-Unhealth\\Health'
        files1 = os.listdir(LN_pic_path_L) + os.listdir(HN_pic_path_H) + os.listdir(NN_pic_path_N1) + os.listdir(
            HL_pic_path_L)
        LN_pic_path_N = '.\\Pic\\Exp_Pic\\Block2\\Health-Nofood\\Nofood'
        HN_pic_path_N = '.\\Pic\\Exp_Pic\\Block2\\Unhealth-Nofood\\Nofood'
        NN_pic_path_N2 = '.\\Pic\\Exp_Pic\\Block2\\Nofood-Nofood\\Nofood2'
        HL_pic_path_H = '.\\Pic\\Exp_Pic\\Block2\\Health-Unhealth\\Unhealth'
        files2 = os.listdir(LN_pic_path_N) + os.listdir(HN_pic_path_N) + os.listdir(NN_pic_path_N2) + os.listdir(
            HL_pic_path_H)
    elif block == 3:
        LN_pic_path_L = '.\\Pic\\Exp_Pic\\Block3\\Health-Nofood\\Health'
        HN_pic_path_H = '.\\Pic\\Exp_Pic\\Block3\\Unhealth-Nofood\\Unhealth'
        NN_pic_path_N1 = '.\\Pic\\Exp_Pic\\Block3\\Nofood-Nofood\\Nofood1'
        HL_pic_path_L = '.\\Pic\\Exp_Pic\\Block3\\Health-Unhealth\\Health'
        files1 = os.listdir(LN_pic_path_L) + os.listdir(HN_pic_path_H) + os.listdir(NN_pic_path_N1) + os.listdir(
            HL_pic_path_L)
        LN_pic_path_N = '.\\Pic\\Exp_Pic\\Block3\\Health-Nofood\\Nofood'
        HN_pic_path_N = '.\\Pic\\Exp_Pic\\Block3\\Unhealth-Nofood\\Nofood'
        NN_pic_path_N2 = '.\\Pic\\Exp_Pic\\Block3\\Nofood-Nofood\\Nofood2'
        HL_pic_path_H = '.\\Pic\\Exp_Pic\\Block3\\Health-Unhealth\\Unhealth'
        files2 = os.listdir(LN_pic_path_N) + os.listdir(HN_pic_path_N) + os.listdir(NN_pic_path_N2) + os.listdir(
            HL_pic_path_H)

    positions_pic = list(range(0, dotexp_picnum))  # 储存不同对图片呈现的顺序，并随机化
    random.shuffle(positions_pic)

    positions_dot = [1] * dotexp_picnum  # 储存点相对于图片的位置，即“一致”或“不一致”，初始化全为一致

    # 接下来的随机化过程有点复杂，主要目的在于将四组LN,HN,NN,HL中每组中两种图片（例如LN中的L和N）的一半与点保持“一致”，提前进行“一致”
    # 或者“不一致”的目的在于提前均衡点的位置与“一致”性
    LN_temp1 = []
    LN_temp2 = []
    HN_temp1 = []
    HN_temp2 = []
    NN_temp1 = []
    NN_temp2 = []
    HL_temp1 = []
    HL_temp2 = []
    for i in range(0, dotexp_picnum):  # 将每组图片位置的一半储存在temp1中，另一半储存在temp2中
        if positions_pic[i] < 12:
            if positions_pic[i] % 2 == 0:
                LN_temp1.append(i)
            else:
                LN_temp2.append(i)
        elif 12 <= positions_pic[i] < 24:
            if positions_pic[i] % 2 == 0:
                HN_temp1.append(i)
            else:
                HN_temp2.append(i)
        elif 24 <= positions_pic[i] < 36:
            if positions_pic[i] % 2 == 0:
                NN_temp1.append(i)
            else:
                NN_temp2.append(i)
        else:
            if positions_pic[i] % 2 == 0:
                HL_temp1.append(i)
            else:
                HL_temp2.append(i)

    random.shuffle(LN_temp1)
    random.shuffle(LN_temp2)
    random.shuffle(HN_temp1)
    random.shuffle(HN_temp2)
    random.shuffle(NN_temp1)
    random.shuffle(NN_temp2)
    random.shuffle(HL_temp1)
    random.shuffle(HL_temp2)
    for i in range(0, int(dotexp_picnum/4/2/2)):  # 使1/4的点与一侧图片不一致，1/4的点与另一侧图片不一致
        positions_dot[LN_temp1[i]] = 2
        positions_dot[LN_temp2[i]] = 2
        positions_dot[HN_temp1[i]] = 2
        positions_dot[HN_temp2[i]] = 2
        positions_dot[NN_temp1[i]] = 2
        positions_dot[NN_temp2[i]] = 2
        positions_dot[HL_temp1[i]] = 2
        positions_dot[HL_temp2[i]] = 2

    timer = core.Clock()  # 设置计时器，用于记录反应时
    for i in range(0, dotexp_picnum):
        fixation = visual.TextStim(win, text='+', pos=[0, 0], height=0.3, color=(-1, -1, -1))
        fixation.draw()
        win.flip()
        clk1 = clock.CountdownTimer(random.uniform(fixation_duration[0], fixation_duration[1]))  # 使注视点随机呈现500-1000ms
        while clk1.getTime() > 0:
            continue

        if positions_pic[i] < 12:  # 使不同组中图片的位置均衡
            if positions_pic[i] % 2 == 0:
                pic_left_path = LN_pic_path_L + '\\' + files1[positions_pic[i]]
                pic_right_path = LN_pic_path_N + '\\' + files2[positions_pic[i]]
            else:
                pic_right_path = LN_pic_path_L + '\\' + files1[positions_pic[i]]
                pic_left_path = LN_pic_path_N + '\\' + files2[positions_pic[i]]
        elif 12 <= positions_pic[i] < 24:
            if positions_pic[i] % 2 == 0:
                pic_left_path = HN_pic_path_H + '\\' + files1[positions_pic[i]]
                pic_right_path = HN_pic_path_N + '\\' + files2[positions_pic[i]]
            else:
                pic_right_path = HN_pic_path_H + '\\' + files1[positions_pic[i]]
                pic_left_path = HN_pic_path_N + '\\' + files2[positions_pic[i]]
        elif 24 <= positions_pic[i] < 36:
            if positions_pic[i] % 2 == 0:
                pic_left_path = NN_pic_path_N1 + '\\' + files1[positions_pic[i]]
                pic_right_path = NN_pic_path_N2 + '\\' + files2[positions_pic[i]]
            else:
                pic_right_path = NN_pic_path_N1 + '\\' + files1[positions_pic[i]]
                pic_left_path = NN_pic_path_N2 + '\\' + files2[positions_pic[i]]
        else:
            if positions_pic[i] % 2 == 0:
                pic_left_path = HL_pic_path_L + '\\' + files1[positions_pic[i]]
                pic_right_path = HL_pic_path_H + '\\' + files2[positions_pic[i]]
            else:
                pic_right_path = HL_pic_path_L + '\\' + files1[positions_pic[i]]
                pic_left_path = HL_pic_path_H + '\\' + files2[positions_pic[i]]

        pic_left = visual.ImageStim(win, image=pic_left_path, size=[1, 1], pos=[-0.5, 0])
        pic_right = visual.ImageStim(win, image=pic_right_path, size=[1, 1], pos=[0.5, 0])
        pic_left.draw()
        pic_right.draw()
        fixation.draw()
        win.flip()

        clk2 = clock.CountdownTimer(1)
        while clk2.getTime() > 0:
            continue

        if positions_dot[i] == 1:  # “一致”时，判断点的位置在左还是右
            if positions_pic[i] % 2 == 0:
                dotposition = -0.5
            else:
                dotposition = 0.5
        else:
            if positions_pic[i] % 2 == 0:
                dotposition = 0.5
            else:
                dotposition = -0.5
        dot = visual.TextStim(win, text='·', pos=[dotposition, 0], height=0.5, color=(-1, -1, -1))
        dot.draw()
        fixation.draw()
        win.flip()

        # keydown = []
        timer.reset()
        # clk3 = clock.CountdownTimer(2)  # 倒计时2s,使点呈现2000ms
        # while clk3.getTime() > 0:  # 键盘检测
        #     if len(keydown) > 0:
        #         continue
        #     else:
        #         keydown = event.getKeys(keyList=['f', 'j'], timeStamped=timer)
        keydown = event.waitKeys(keyList=['f', 'j'])
        time_press = timer.getTime()
        # print(keydown)

        if (keydown[0] == 'f' and dotposition == -0.5) or (keydown[0] == 'j' and dotposition == 0.5):
            dotexp_result[i] = 1  # 被试按键正确则记为1

        reaction_time[i] = time_press
        pic_pair_class[i] = int(positions_pic[i] / 12)

    return dotexp_result, reaction_time, positions_dot, pic_pair_class


def input_method_change():
    """
    参考https://blog.csdn.net/jpch89/article/details/84281136
    :return:
    """
    # 语言代码
    # https://msdn.microsoft.com/en-us/library/cc233982.aspx
    LID = {0x0804: "Chinese (Simplified) (People's Republic of China)",
           0x0409: 'English (United States)'}

    # 获取前景窗口句柄
    hwnd = win32gui.GetForegroundWindow()

    # 获取前景窗口标题
    title = win32gui.GetWindowText(hwnd)
    # print('当前窗口：' + title)

    # 获取键盘布局列表
    im_list = win32api.GetKeyboardLayoutList()
    im_list = list(map(hex, im_list))
    # print(im_list)

    # 设置键盘布局为英文
    result = win32api.SendMessage(
        hwnd,
        WM_INPUTLANGCHANGEREQUEST,
        0,
        0x0409)
    if result == 0:
        print('设置英文键盘成功！')


if __name__ == '__main__':

    app = QApplication(sys.argv)
    subjectInformation = subInfoWindow()
    subjectInformation.show()  # 呈现被试信息填写界面
    app.exec_()  # 将exit()加入主循环直至被调用

    win = visual.Window(pos=[0.5, 0.5], color=(1, 1, 1), fullscr=True)
    intro_img_show(win, imgPath='.\\pic\\intro\\Pra_Intro.png')

    intro_img_show(win, imgPath='.\\pic\\intro\\Pra_MIST_Intro.png')

    MIST = MISTWindow(block=0, group=0)  # 练习阶段的MIST任务
    MIST.show()
    app.exec_()
    MeanT = MIST.mean_time
    print('平均答题时间为：{}s'.format(MeanT))

    intro_img_show(win, imgPath='.\\pic\\intro\\Pra_Dot_Intro.png')
    # 不断进行练习阶段的点侦测任务，直到正确率大于correctRate
    input_method_change()  # 改变输入法至英文
    while 1:
        # main_window_state(win, 1)
        cRate = pra_dotexp(win)
        if cRate > correctRate:
            break
        else:
            intro_img_show(win, imgPath='.\\pic\\intro\\Pra_Dot_Intro_2.png')

    intro_img_show(win, imgPath='.\\pic\\intro\\Pra_End.png')  # 练习阶段结束

    block_seq = []  # 按顺序不同储存block序号
    if subinfo[2] == 1:
        block_seq = [1, 2, 3]
    elif subinfo[2] == 2:
        block_seq = [2, 1, 3]

    dotexp_results = numpy.zeros((dotexp_picnum, block_num))  # 初始化数据储存变量
    RTs = numpy.zeros((dotexp_picnum, block_num))
    positionDots = numpy.zeros((dotexp_picnum, block_num))
    pair_pic_classes = numpy.zeros((dotexp_picnum, block_num))

    for i in range(block_num):  # 不同block的循环

        # main_window_state(win, 2)  # 调整psychopy打开窗口的优先级
        MIST = MISTWindow(block=block_seq[i], group=subinfo[1])  # MIST任务呈现
        MIST.show()
        app.exec_()

        # main_window_state(win, 1)
        # input_method_change()
        dotexp_result, RT, positionDot, pair_pic_class = dotexp(win, block_seq[i])  # 点侦测任务，得到被试任务相关数据并处理
        dotexp_results[:, i] = numpy.array(dotexp_result)
        RTs[:, i] = numpy.array(RT)
        positionDots[:, i] = numpy.array(positionDot)
        pair_pic_classes[:, i] = numpy.array(pair_pic_class)
        if i < block_num-1:
            rest_img = visual.ImageStim(win, image='.\\pic\\intro\\Rest.png', size=[1.5, 1.5])  # 呈现休息图片
            rest_img.draw()
            win.flip()
            # core.wait(300)  # 填问卷加休息共设为5min
            intro_img_show(win, imgPath='.\\pic\\intro\\rest02.png')  # 呈现休息图片

    end_img = visual.ImageStim(win, image='.\\pic\\intro\\Exp_End.png', size=[1.5, 1.5])  # 呈现结束图片，等待3s退出
    end_img.draw()
    win.flip()
    core.wait(3)

    # print(dotexp_results)
    # print(RTs)
    # print(positionDots)

    resultsbook_name = '.\\data\\' + subinfo[0] + '.xlsx'  # 储存结果
    resultsbook = xlsxwriter.Workbook(resultsbook_name)
    sheet = resultsbook.add_worksheet('sheet1')
    sheet.write(0, 0, 'ID')
    sheet.write(0, 1, '组别')
    sheet.write(0, 2, '顺序')
    sheet.write(0, 3, '性别')
    sheet.write(0, 4, '年龄')
    sheet.write(0, 5, 'Block')
    sheet.write(0, 6, 'Trial Num')
    sheet.write(0, 7, '反应时')
    sheet.write(0, 8, '是否正确（1正确0，不正确）')
    sheet.write(0, 9, '是否一致（1一致 2不一致）')
    sheet.write(0, 10, '图片种类（0=LN, 1=HN, 2=NN, 3=HL）')

    for i in range(1, dotexp_picnum*block_num+1):
        sheet.write(i, 0, subinfo[0])
        sheet.write(i, 1, subinfo[1])
        sheet.write(i, 2, subinfo[2])
        sheet.write(i, 3, subinfo[3])
        sheet.write(i, 4, subinfo[4])
        if i <= dotexp_picnum:
            sheet.write(i, 5, 1)
        elif i <= dotexp_picnum*2:
            sheet.write(i, 5, 2)
        else:
            sheet.write(i, 5, 3)
        sheet.write(i, 6, (i-1) % dotexp_picnum + 1)  # 记录试次
        sheet.write(i, 7, RTs[i%dotexp_picnum-1, int((i-1)/dotexp_picnum)])
        sheet.write(i, 8, dotexp_results[i%dotexp_picnum-1, int((i-1)/dotexp_picnum)])
        sheet.write(i, 9, positionDots[i%dotexp_picnum-1, int((i-1)/dotexp_picnum)])
        sheet.write(i, 10, pair_pic_classes[i%dotexp_picnum-1, int((i-1)/dotexp_picnum)])

    resultsbook.close()
    print('被试数据已储存')
