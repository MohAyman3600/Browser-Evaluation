from selenium import webdriver
import platform
try:
        from PIL import ImageGrab
except:
        import pyscreenshot as ImageGrab
import xlsxwriter
import subprocess #To get the resolution
import os
import sys
import re
try:
        from win32api import GetSystemMetrics
except:
		pass

currentDirPath = os.getcwd()
 

with open("links.txt") as f:
        content = f.readlines()
links = ['http://'+x.strip() for x in content]

def openFirefox(url,i):
        browser = webdriver.Firefox()
        browser.get(url)
        browser.maximize_window()
        domain = browser.current_url
        occur = domain.count('xn--')
        start = "http://xn--"
        if occur <= 1:
                end = "."
        else:
                end = ".xn--"
        end2 = "/"
        result = domain[domain.find(start)+len(start):domain.rfind(end)]
        result = result.decode('punycode')
        result2 = domain[domain.find(end)+len(end):domain.rfind(end2)]
        if occur > 1:
                result2 = result2.decode('punycode')
        else:
                result2 = result2
        m = "http://"
        urlBefore = url[url.find(m)+len(m):].decode("utf-8","ignore")
        urlAfter = result+"."+result2
        if( urlBefore == urlAfter):
                condition = "Match"
        else:
                condition = "Does not Match"
        worksheet.write_url('A'+str(cellfir+1), urlBefore,format)
        worksheet.write_url('A'+str(cellfir+2), urlAfter,format)
        worksheet.write_url('A'+str(cellfir+4), 'Condtion : '+condition)
        im=ImageGrab.grab()
        # part of the screen
        crop_width = (int(width)*80)/100
        crop_height = (int(height)*15)/100
        im=ImageGrab.grab(bbox=(0,0,crop_width,crop_height))
        # # im.show()
        # # to file
        im.save('screenfir'+str(i)+'.png')
        browser.quit()


def openChrome(url,i):
        try:
                browser = webdriver.Chrome(currentDirPath+'/chromedriver.exe')
        except:
                browser = webdriver.Chrome(currentDirPath+'/chromedriver')
        browser.get(url)
        browser.maximize_window()
        domain = browser.current_url
        occur = domain.count('xn--')
        start = "http://xn--"
        if occur <= 1:
                end = "."
        else:
                end = ".xn--"
        end2 = "/"
        result = domain[domain.find(start)+len(start):domain.rfind(end)]
        result = result.decode('punycode')
        result2 = domain[domain.find(end)+len(end):domain.rfind(end2)]
        if occur > 1:
                result2 = result2.decode('punycode')
        else:
                result2 = result2
        m = "http://"
        urlBefore = url[url.find(m)+len(m):].decode("utf-8","ignore")
        urlAfter = result+"."+result2
        if( urlBefore == urlAfter):
                condition = "Match"
        else:
                condition = "Does not Match"
        worksheet.write_url('A'+str(cellchr+1), urlBefore,format)
        worksheet.write_url('A'+str(cellchr+2), urlAfter,format)
        worksheet.write_url('A'+str(cellchr+4), 'Condtion : '+condition)
        im=ImageGrab.grab()
        crop_width = (int(width)*80)/100
        crop_height = (int(height)*15)/100
        im=ImageGrab.grab(bbox=(0,0,crop_width,crop_height))
##        im.show()
        im.save('screenchr'+str(i)+'.png')
        browser.quit()

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('Monitoring.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
format = workbook.add_format()

format.set_align('center_across')
format.set_align('vcenter')

worksheet.set_column('A:A', 30)

if platform.system() == 'Linux':
        cmd = ['xrandr']
        cmd2 = ['grep', '*']
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE)
        p2 = subprocess.Popen(cmd2, stdin=p.stdout, stdout=subprocess.PIPE)
        p.stdout.close()

        resolution_string, junk = p2.communicate()
        resolution = resolution_string.split()[0]
        width, height = resolution.split('x')

        worksheet.write('A1','Linux')
        cellfir = 2
        cellchr = 10
        for x in range(0,len(links)):
                url = links[x]
                openFirefox(url,x)
                # Insert an image.
                worksheet.write('A'+str(cellfir), 'Firefox')
                worksheet.insert_image('B'+str(cellfir), 'screenfir'+str(x)+'.png',{'x_offset': 15, 'y_offset': 10,'x_scale': 0.5, 'y_scale': 0.5})
                worksheet.write_url('B'+str(cellfir+6), currentDirPath+'\screenfir'+str(x))
                openChrome(url,x)
                # Insert an image.
                worksheet.write('A'+str(cellchr), 'Chrome')
                worksheet.insert_image('B'+str(cellchr), 'screenchr'+str(x)+'.png',{'x_offset': 15, 'y_offset': 10,'x_scale': 0.5, 'y_scale': 0.5})
                worksheet.write_url('B'+str(cellchr+6), currentDirPath+'\screenchr'+str(x))
                cellfir += 15
                cellchr += 15

if platform.system() == 'Windows':
        width = GetSystemMetrics(0)
        height = GetSystemMetrics(1)
        worksheet.write('A1','Windows')
        cellfir = 2
        cellchr = 10
        for x in range(0,len(links)):
                url = links[x]
                openFirefox(url,x)
                # Insert an image.
                worksheet.write('A'+str(cellfir), 'Firefox')
                worksheet.insert_image('B'+str(cellfir), 'screenfir'+str(x)+'.png',{'x_offset': 15, 'y_offset': 10,'x_scale': 0.5, 'y_scale': 0.5})
                worksheet.write_url('B'+str(cellfir+6), currentDirPath+'\screenfir'+str(x))
                openChrome(url,x)
                # Insert an image.
                worksheet.write('A'+str(cellchr), 'Chrome')
                worksheet.insert_image('B'+str(cellchr), 'screenchr'+str(x)+'.png',{'x_offset': 15, 'y_offset': 10,'x_scale': 0.5, 'y_scale': 0.5})
                worksheet.write_url('B'+str(cellchr+6), currentDirPath+'\screenchr'+str(x))
                cellfir += 15
                cellchr += 15
        
reload(sys)  
sys.setdefaultencoding('utf8')

workbook.close()
