import tkinter as tk  
import tkinter.ttk as ttk          #Comboboxを使うため
from tkcalendar import DateEntry   #外部ライブラリのtkcalendarモジュールのDateEntryメソッド    【 pip install tkcalendar 】
from tkinter import messagebox
import sqlite3
import csv
#fromで読み込んだモジュールのメソッドは,モジュール名をつけることなく使用できる.例:messagebox.showinfo････
#from tkinter import *
    
       
#DB接続
conn = sqlite3.connect('incident.db')    #DBファイルを指定する
cur = conn.cursor()
#tableの作成    IF NOT EXISTS は一度作成していたら作成はしないよってこと.
sql_create = '''CREATE TABLE IF NOT EXISTS incidenttable(date,client,staff,item,support )'''
cur.execute(sql_create) #('''CREATE TABLE IF NOT EXISTS incidenttable(date,client,staff,item,support )''')
conn.commit()
conn.close()


 
    
    
#子クラスを定義,親から継承        
class Application(tk.Frame): 
     # (master=root=tk.Tk())  masterは空ですよってこと.
    def __init__(self,master=None): 
        #ここでいうsuper().__init__()は,tk.Frameと同じ意味
        super().__init__(master)  #frame=tk.Frame(root,･･･)       外からmasterを受け取って基底クラスのコンストラクタに渡す 受け取りは上の段のmaster   基底クラスのコンストラクタをオーバーライド,super().__init__をつけることで,親クラスの__init__と同じ処理ができる.
        #インスタンス変数
         #クラスの他のメソッドで使うためインスタンス変数に代入しておく
        #self.root = master                
        self.master = master         #master = root = tk.Tk()     
        self.master.title('インシデント管理システム')
        self.master.geometry('500x370')
        
        #Applicationクラスのselfは,Frameのことだと思ってよい.
        #(fill='both')   #Frameの位置を定義   /packオプション  【fill】  x 横拡大   y 縦拡大  both 縦横拡大/
        self.pack() 
        self.configure(width=2500,height= 1500,bg='lightgrey')
        self.propagate(False)  
        #create_widgetsメソッドを呼び出す
        self.create_widgets() 
    


     #インスタンスメソッド
     #各ウィジェッツ機能
     #クラス内のメソッドには(self)とつける.
    def create_widgets(self):   
    
        #メニューラベル
        #このselfは,Frameのこと.
        menu_label = tk.Label(self)  
        menu_label['text'] = 'メニュー'
        menu_label['relief'] = tk.RAISED
        #変数menu_fontを参照
        menu_label['font'] = menu_font   
        menu_label['bd'] = 8
        menu_label['width'] =20
        menu_label['height'] = 1
        menu_label.pack(pady=50)   
       
        #インシデント入力ボタン
        input_button = tk.Button(self)
        input_button['text'] ='入力処理'
        input_button['font'] = font
        input_button['width'] = 30
        #sub_windowメソッドを呼び出す.他の関数を呼び出している.
        input_button['command'] = self.sub_window1 
        #input_button['fg']    = 'red'
        #input_button['activeforeground']="red"
        input_button.pack()
        
        #照会処理ボタン
        reference_button = tk.Button(self)
        reference_button['text'] = '照会処理'
        reference_button['font'] = font
        reference_button['width'] = 30
        reference_button['command'] = self.sub_window2
        reference_button.pack(pady=25)
        
        #CSV出力ボタン
        output_button = tk.Button(self)
        output_button['text'] = 'CSV出力'
        output_button['font'] = font
        output_button['width'] =30
        output_button['command'] = self.CSV_output
        output_button.pack()
        
        #終了ボタン
        quit_button = tk.Button(self)
        quit_button['text'] = '閉じる'
        quit_button['font'] = font
        quit_button['width'] = 15
        #self.root(=tk.Tk())を呼び出して,destroyを使用
        quit_button['command'] =self.master.destroy  
        quit_button.pack(side='bottom',pady=5)
        
        
        
        
    #サブウィンドウ機能 (ボタンを押して呼び起こすようにする)
    def sub_window1(self):
        sub_win1 = tk.Toplevel(self)
        sub_win1.title('入力画面')
        sub_win1.geometry('700x450')
        sub_frame1 = tk.Frame(sub_win1,width=2500,height= 1500,bg='lightgrey')
        sub_frame1.pack()#(fill='both')
        
        def  import_SQL():
            #入力値を取得
            Value1  = sub1_dayentry.get_date()
            Value2 = sub1_entry2.get()
            Value3  =sub1_entry3.get()
            Value4  =sub1_combobox.get()
            #引数:最初から最後までという意味
            Value5 =sub1_text5.get("1.0",'end')   
            
            conn = sqlite3.connect('incident.db')
            cur = conn.cursor()
            
            sql_insert = 'insert into incidenttable values(?,?,?,?,?)'
            data = [(Value1,Value2,Value3,Value4,Value5),]
            try:
                # cur.executemany('insert into incidenttable values(?,?,?,?,?)',data)
                cur.executemany(sql_insert,data)           
                conn.commit()
                messagebox.showinfo('登録確認','登録完了しました。') #tkinterのmessageboxメソッド
                #for i in data:
                 #   messagebox.showinfo('登録完了','{0}を入力しました。'.format(i))
                conn.close()
            except:
                pass
    
        
        #年月日ラベル･エントリ
        sub1_label1 = tk.Label(sub_frame1)
        sub1_label1['text'] ='年 月 日'
        sub1_label1['width'] = 20
        sub1_label1['relief'] = 'sunken'
        sub1_label1['font']  = font
        sub1_label1.place(x=45,y=50)
        #sub1_label1.grid(row=1,column=0,columnspan=2)
        #sub1_label1.pack(side='left')
        
        #日付カレンダ入力
        sub1_dayentry = DateEntry(sub_frame1,selectmode='day',showweeknumbers=False)
        sub1_dayentry.place(x=250,y=50)
        Value1 = sub1_dayentry.get_date()
        
        """
        sub1_entry1 = tk.Entry(sub_win1)
        sub1_entry1['width'] = 30           #文字入力数(半角)
        sub1_entry1['font'] = font
        sub1_entry1['justify'] = tk.LEFT    #その他,CENTER,RIGHT
        sub1_entry1.place(x=250,y=50)
        #sub1_entry1.grid(row=1,column=2,)
        #sub1_entry.pack(side='left')
        """
        
        #顧客名ラベル･エントリ
        sub1_label2 = tk.Label(sub_frame1)
        sub1_label2['text'] ='顧 客 名'
        sub1_label2['width'] = 20
        sub1_label2['font']  = font
        sub1_label2['relief'] = 'sunken'
        sub1_label2.place(x=45,y=100)
        #sub1_label2.grid(row=2,column=0)
        sub1_entry2 = tk.Entry(sub_frame1)
        sub1_entry2['width'] = 30
        sub1_entry2['font'] = font
        sub1_entry2['justify'] = tk.LEFT
        sub1_entry2.place(x=250,y=100)
        #sub1_entry2.grid(row=2,column=2)
        Value2 = sub1_entry2.get()
        
        
        #担当者ラベル･エントリ
        sub1_label3 = tk.Label(sub_frame1)
        sub1_label3['text'] ='担 当 者'
        sub1_label3['width'] = 20
        sub1_label3['font']  = font
        sub1_label3['relief'] = 'sunken'
        sub1_label3.place(x=45,y=150)
        #sub1_label3.grid(row=2,column=0)
        sub1_entry3 = tk.Entry(sub_frame1)
        sub1_entry3['width'] = 30
        sub1_entry3['font'] = font
        sub1_entry3['justify'] = tk.LEFT
        sub1_entry3.place(x=250,y=150)
        #sub1_entry3.grid(row=2,column=2)
        Value3 = sub1_entry3.get()
        
  
        #トラブル項目ラベル・コンボボックス
        sub1_label4 = tk.Label(sub_frame1)
        sub1_label4['text'] = 'トラブル項目'
        sub1_label4['width'] = 20
        sub1_label4['font'] = font
        sub1_label4['relief'] = 'sunken'
        sub1_label4.place(x=45,y=200)
        #コンボボックス
        module =('ハードウェア','ソフトウェア','ネットワーク','セキュリティ','データベース','その他')
        #親ウィジェットまたはメインウィンドウを指定
        sub1_combobox = ttk.Combobox(sub_frame1) 
        sub1_combobox['values'] = module
        #ここでいう高さとは,ドロップダウンリストで表示するデータ件数のこと
        sub1_combobox['height'] = 4 
        sub1_combobox['width'] =28
        sub1_combobox['justify'] = 'left'
        sub1_combobox['font'] = font
        sub1_combobox.place(x=250,y=200) 
        Value4 = sub1_combobox.get()
        
        
        #トラブル項目エントリ
        """
        sub1_entry4 = tk.Entry(sub_frame1)
        sub1_entry4['width'] = 30
        sub1_entry4['font'] = font
        sub1_entry4['justify'] =tk.LEFT
        #sub1_entry4.insert(tk.END,'トラブル内容を入力')   #事前に入力されている値
        sub1_entry4.place(x=250,y=200)
        """
        
        
        #対応内容ラベル
        sub1_label5=  tk.Label(sub_frame1)
        sub1_label5['text'] = '対応内容'
        sub1_label5['width'] = 20
        sub1_label5['font'] = font
        sub1_label5['relief'] = 'sunken'
        sub1_label5.place(x=45,y=250)
        #対応内容Text
        sub1_text5 = tk.Text(sub_frame1)
        sub1_text5['width'] =  38
        sub1_text5['height'] = 5
        #sub1_text5.configure(font=("Courier", 16, "italic"))
        #sub1_text5.insert(tk.END,'対応内容を入力')   #事前に入力されている値
        sub1_text5.place(x=250,y=250)
        sub1_text5.configure(font=font) #configure()メソッドはウィンドウやウィジェットの属性(オプション)を変更するときに使用
        Value5 = sub1_text5.get('1.0','end')
        
        
        #DB登録ボタン
        button1 = tk.Button(sub_frame1)
        button1['text'] = '登録'
        button1['command'] = import_SQL
        button1['font'] =font
        button1['width'] = 12
        #button1.pack(side='bottom',pady=10)
        button1.place(x=300,y=360)
        #sub_win1.mainloop()
        
        
    
    def sub_window2(self):
        sub_win2 = tk.Toplevel(self)
        sub_win2.title('表示画面')
        sub_win2.geometry('700x670')
        #sub_win2.configure(bg='lightgrey')  #configure()メソッドはウィンドウやウィジェットの属性（オプション）を変更することに使用されます
        sub_frame2 = tk.Frame(sub_win2,bg='lightgrey')   #(,width=2500,height=1500)
        # sub_frame2.configure(bg='lightgrey')
        sub_frame2.pack()
        
       #期間を指定して,DBからデータを取得
        def selct_SQL():
            # for i in tree.get_children():
            #     tree.delete(i)
            tree.delete(*tree.get_children())
            start = day_entry1.get().replace('/','-')
            end   = day_entry2.get().replace('/','-')
    
            if start == '':
               start = '1900-01-01'
            if end   == '':
                end = '2100-01-01'
        
            #DB接続
            conn = sqlite3.connect('incident.db')
            cur = conn.cursor()
            #クエリ select文
            # cur.execute('''SELECT * FROM incidenttable WHERE date BETWEEN  '{}' AND '{}' ORDER  BY   date'''.format(start,end))
            # cur.execute(f'''SELECT * FROM incidenttable WHERE date BETWEEN  '{start}' AND '{end}' ORDER  BY   date''')  #f文字列を使用
            sql_select = f'''SELECT * FROM incidenttable WHERE date BETWEEN  '{start}' AND '{end}' ORDER  BY   date'''
            cur.execute(sql_select)
            #sqlite3の場合,日付型(date型)がない.
            #where date BETWEEN {}AND{}
    

            for i in cur:
                #ツリービューの要素に追加
                tree.insert('','end',values=i)
    
            conn.close()
    
        #表示画面ラベル
        display_label = tk.Label(sub_frame2)
        display_label['text'] = '表 示 画 面'
        display_label['font'] =('',20,'bold')
        display_label['width'] = 10
        display_label['height'] = 1
        display_label['relief'] = 'ridge'
        display_label['bg'] = 'lavender'
        display_label.pack(pady=10)


        #フレームの作成(期間選択のラベルとエントリーの設定)
        frame_labelentry = tk.Frame(sub_frame2,bg='lightgrey')
        frame_labelentry.pack()   #bd=2,relief='ridge'

        day_label = tk.Label(frame_labelentry)
        day_label['font'] = ('',14)
        day_label['text'] = '期間'
        day_label.pack(side='left')

        day_entry1 = tk.Entry(frame_labelentry)
        day_entry1['font'] =('',14)
        day_entry1['justify'] = 'center'
        day_entry1['width'] = 12
        day_entry1.pack(side = 'left')
        day_entry1.get()

        day_label2 = tk.Label(frame_labelentry)
        day_label2['font'] = ('',14)
        day_label2['text'] = '～'
        day_label2.pack(side='left')

        day_entry2 = tk.Entry(frame_labelentry)
        day_entry2['font'] =('',14)
        day_entry2['justify'] = 'center'
        day_entry2['width'] = 12
        day_entry2.pack(side = 'left')
        day_entry2.get()
        
        #表示ボタン
        display_button = tk.Button(frame_labelentry)
        display_button['text'] ='表示'
        display_button['font'] = ('',14)
        display_button['width'] =12
        display_button['bg'] = 'gray'
        display_button['command'] = selct_SQL
        display_button.pack(side='left')



        #Treeview
        #ツリービューの作成
        tree = ttk.Treeview(sub_frame2,padding=10)
        #列インデックスの作成  
        tree['columns'] = (1,2,3,4,5)
        #表スタイルの設定(headingsはツリー形式ではない、通常の表形式)  
        tree['show'] ='headings'
        #表示件数  簡単に言うと行の数  
        tree['height'] = 10 


        #ttkモジュールの多くは,【縦幅】をStyleを使って設定(テーマを変更している場合,影響をうけるかもしれない.)
        style = ttk.Style()
        #Treeviewの全部
        style.configure('Treeview',rowheight=60)     
        #TreeviewのHeading部分のみ    
        style.configure("Treeview.Heading",font=("",12,'bold'))  
        # 全てのウィジェット
        #style.configure(".",font=("",14))                


        #各列の設定(インデックス番号,オプション(今回は番号))  各columnの横幅はここで設定
        tree.column(1,width=110)
        tree.column(2,width=110)
        tree.column(3,width=110)
        tree.column(4,width=110)
        tree.column(5,width=200 )

        #各列のヘッダー設定(インデックス番号,テキスト名)
        tree.heading(1,text ='日付')
        tree.heading(2,text='顧客名')
        tree.heading(3,text='担当者')
        tree.heading(4,text='トラブル項目')
        tree.heading(5,text='対応内容')
        
        #ツリービューのは位置
        tree.pack(fill='x',padx=20,pady=20)

        
        
        
    #CSV出力機能
    def CSV_output(self):
        conn = sqlite3.connect('incident.db')
        cur =conn.cursor()
        cur.execute('select * from incidenttable')

        with open ('C:/Users/socce/Documents/test.csv',mode='w',newline='')as f:
            csv_writer =csv.writer(f)
            #  cur.description:列の項目名を出力する.   i[0]とは変数iの一列目を表示すること
            csv_writer.writerow( [ i[0] for i in cur.description]) 
            csv_writer.writerows(cur)
        
        messagebox.showinfo('CSV出力','出力完了しました｡')
        
        conn.close()
        
        
        
        """
        sub_win3 = tk.Toplevel(self)
        sub_win3.title('subwindow3')
        sub_win3.geometry('500x370')
        sub_frame3 = tk.Frame(sub_win3,width=2500,height= 1500,bg='lightgrey')
        sub_frame3.pack(fill='both')     """
        



#フォント用変数
#メニュ用フォント
menu_font = ('MS明朝',18,'bold')  
#'bold':太文字  'italic':斜め   
font = ('MS明朝',13)     
        
              
if __name__  ==  '__main__':             #このファイルが実行されている場合の処理
    #インスタンスを生成
    root = tk.Tk()                       #rootインスタンスを生成
    app = Application( master = root )   #appインスタンスを生成    rootをmasterに代入  引数としてクラスに代入
    app.mainloop()                       #appインスタンスのイベントハンドラを呼び出し    


