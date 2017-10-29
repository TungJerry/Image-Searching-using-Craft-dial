from cStringIO import StringIO
from PIL import Image
from icrawler.builtin import GoogleImageCrawler
from Craft import CraftClient
import time
import sys
import os
import win32clipboard
import win32com.client
import shutil

PATH_TO_CRAFT_SDK_FILE = "./"
sys.path.insert(0, PATH_TO_CRAFT_SDK_FILE)

pre_recv_data =''
recv_data = ''
pre_search_idx = 0
now_search_idx = 0
result_link = []
download_finish = False
start_download = False
counter = 0
shell = win32com.client.Dispatch("WScript.Shell")


def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def recv_from_clipboard(clip_type):
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return data

def load_image(image_idx):
    global shell
    
    filepath = ""
    if image_idx == "waiting":
        filepath = 'images/' + image_idx + '.jpg'
    else:
        for filename in os.listdir("./images/google"):
            #print "in download dir: " + filename
            if filename.split('.')[0] == image_idx:
                print "match the file!!!"
                if filename.endswith(".png"):
                    print('its extension is png')
                    filepath = 'images/google/' + image_idx + '.png'
                elif filename.endswith(".gif"):
                    print('its extension is gif')
                    filepath = 'images/google/' + image_idx + '.gif'
                else:
                    print('its extension is jpg')
                    filepath = 'images/google/' + image_idx + '.jpg'
                break

    print  "God damn sucking filepath :" + filepath + "\n"
    image = Image.open(filepath)

    output = StringIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)
    time.sleep(0.03)
    shell.SendKeys("^v")

def copy_word_and_download_image():
    global recv_data
    global shell

    shell.SendKeys("^c")
    time.sleep(0.05)
    recv_data = recv_from_clipboard(win32clipboard.CF_DIB)
    shell.SendKeys("{DELETE}")
    load_image('waiting')

    google_crawler = GoogleImageCrawler(
        downloader_threads=8,
        storage={'root_dir': 'images/google'}
        )
    google_crawler.crawl(
        str(recv_data) ,
        max_num=8,
        date_min=None,
        date_max=None,
        min_size=(200,200),
        max_size=(800,800))

def handleCraftEvent_chrome_gmail(event):
    # do tap
    global counter 
    counter = 0
    global pre_search_idx
    global now_search_idx
    global download_finish
    global start_download
    global shell

    #print '\n\n' + str(event) + '\n\n'



    # do turn
    if (event['message_type'] == 'crown_turn_event'):
        #check tap first
        print 'in turn state\n'
        if event['delta'] < 0:
            if(now_search_idx > 10) and download_finish:
                now_search_idx -= 1
        elif event['delta'] > 0:
            if not download_finish and not start_download:
                start_download = True
                try:
                    copy_word_and_download_image()
                except:
                    return
                download_finish = True
                start_download = False
                now_search_idx = 10
                tmp = now_search_idx / 10
                shell.SendKeys("{DELETE}")
                load_image(str(tmp).zfill(6))
                pre_search_idx = tmp
            
            if(now_search_idx < 80) and download_finish:
                now_search_idx += 1

        tmp = now_search_idx / 10
        print 'search idx: %d\n' % tmp
        if(pre_search_idx != tmp) and download_finish:
            shell.SendKeys("{DELETE}")
            load_image(str(tmp).zfill(6))
            pre_search_idx = tmp
        #else:
            #stopMoving()

def handleCraftEvent_ppt(event):
    # do tap
    global counter 
    counter = 0
    global pre_search_idx
    global now_search_idx
    global download_finish
    global start_download
    global shell

    #print '\n\n' + str(event) + '\n\n'



    # do turn
    if (event['message_type'] == 'crown_turn_event'):
        #check tap first
        print 'in turn state\n'
        if event['delta'] < 0:
            if(now_search_idx > 10) and download_finish:
                now_search_idx -= 1
        elif event['delta'] > 0:
            if not download_finish and not start_download:
                start_download = True
                try:
                    copy_word_and_download_image()
                except:
                    return
                download_finish = True
                start_download = False
                now_search_idx = 10
                tmp = now_search_idx / 10
                shell.SendKeys("{BACKSPACE}")
                load_image(str(tmp).zfill(6))
                pre_search_idx = tmp
            
            if(now_search_idx < 80) and download_finish:
                now_search_idx += 1

        tmp = now_search_idx / 10
        print 'search idx: %d\n' % tmp
        if(pre_search_idx != tmp) and download_finish:
            shell.SendKeys("{BACKSPACE}")
            load_image(str(tmp).zfill(6))
            pre_search_idx = tmp
        #else:
            #stopMoving()

def handleCraftEvent_word(event):
    # do tap
    global counter 
    counter = 0
    global pre_search_idx
    global now_search_idx
    global download_finish
    global start_download
    global shell

    #print '\n\n' + str(event) + '\n\n'



    # do turn
    if (event['message_type'] == 'crown_turn_event'):
        #check tap first
        print 'in turn state\n'
        if event['delta'] < 0:
            if(now_search_idx > 10) and download_finish:
                now_search_idx -= 1
        elif event['delta'] > 0:
            if not download_finish and not start_download:
                start_download = True
                try:
                    copy_word_and_download_image()
                except:
                    return
                download_finish = True
                start_download = False
                now_search_idx = 10
                tmp = now_search_idx / 10
                shell.SendKeys("{BACKSPACE}")
                time.sleep(0.01)
                shell.SendKeys("{BACKSPACE}")
                load_image(str(tmp).zfill(6))
                pre_search_idx = tmp
            
            if(now_search_idx < 80) and download_finish:
                now_search_idx += 1

        tmp = now_search_idx / 10
        print 'search idx: %d\n' % tmp
        if(pre_search_idx != tmp) and download_finish:
            shell.SendKeys("{BACKSPACE}")
            time.sleep(0.01)
            shell.SendKeys("{BACKSPACE}")
            load_image(str(tmp).zfill(6))
            pre_search_idx = tmp
        #else:
            #stopMoving()

craft_ppt = CraftClient()
# for windows
craft_ppt.connect("POWERPNT.EXE", "")
craft_ppt.registerEventHandler(handleCraftEvent_ppt)

craft_word = CraftClient()
# for windows
craft_word.connect("WINWORD.EXE", "")
craft_word.registerEventHandler(handleCraftEvent_word)

craft_chrome = CraftClient()
# for windows
craft_chrome.connect("chrome.exe", "")
craft_chrome.registerEventHandler(handleCraftEvent_chrome_gmail)


while(1):
    global counter
    time.sleep(1)
    if download_finish:
        counter+=1
    if counter==5 :
        download_finish = False
        folder = './images/google'
        for the_file in os.listdir(folder):
            file_path = os.path.join(folder, the_file)
            try:
                if os.path.isfile(file_path):
                    print 'delete' + file_path
                    os.unlink(file_path)
            #elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                print(e)

