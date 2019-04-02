# pip install opencv-python
# no need of https://pypi.org/project/Pillow/5.4.1/
# pip install pywin32 for emails

import cv2
import numpy as np
from datetime import datetime as dt
import os
import winsound
from time import sleep
from threading import Thread

initial_timeout = 2
alpha = 0.5
T3 = 1000
cam0 = cv2.VideoCapture(0)
time_between_captures = 200
n_pixels_in_motion = 50
timeout_sending = 10
timeout_counter = 2
n_attachments_per_mail = 3


def get_image(camera):
   # read is the easiest way to get a full image out of a VideoCapture object.7
   _, im = camera.read()
   return cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)


def get_moment():
   return dt.now().strftime('%Y-%m-%d_%H-%S-%f')


def get_image_name(i):
   dir_path = os.getcwd() + '\snapshots'
   if not os.path.exists(dir_path):
     os.makedirs(dir_path)
   os.path.dirname(dir_path)
   return dir_path + '\img_{}_{:d}.jpg'.format(get_moment(),i)
   # return dir_path + '\img_{:d}.jpg'.format(i)

def send_mail(attachments):
   import copy
   at_names = copy.deepcopy(attachments)
   import win32com.client as win32
   outlook = win32.Dispatch('outlook.application')
   mail = outlook.CreateItem(0)
   # mail.To = 'roberto.costa@sas.com'
   mail.To = 'rcode251@gmail.com'
   mail.Subject = 'Check this please'
   mail.Body = 'Do you recognize?'
   mail.HTMLBody = '<h2>See attachment</h2>'  # this field is optional

   # To attach a file to the email (optional):
   for attachment in at_names:
      if os.path.isfile(attachment):
         mail.Attachments.Add(attachment)
   print(mail)
   mail.Send()

def beep_thread(freq_multi,duration_multi):
   duration = 100  # millisecond
   freq = 880  # Hz
   winsound.Beep(freq_multi*freq, duration*duration_multi)

def beep(freq_multi,duration_multi):
   # beep
   thread = Thread(target=beep_thread, args=(freq_multi,duration_multi))
   thread.start()
   return

if __name__ == "__main__":
   # from multiprocessing.pool import ThreadPool
   # pool = ThreadPool(processes=1)

   sleep(initial_timeout)
   img = get_image(cam0)
   cv2.imshow('img', img)
   cv2.waitKey(time_between_captures)
   B = img
   prev_input = img
   motion = np.zeros(img.shape, dtype='uint8')
   started_sending = False
   first_now_sending = dt.now()
   first_now_counter = dt.now()
   attachments_name = []
   i=0

   while True:
      img = get_image(cam0)
      cv2.imshow('img', img)
      cv2.waitKey(time_between_captures)

      B = alpha * prev_input + (1 - alpha) * B
      rho = img-B
      motion = rho**2>T3
      prev_input = img



      if np.count_nonzero(motion) > n_pixels_in_motion:
         im_name = get_image_name(i)
         cv2.imwrite(im_name, img)
         print("movement " + im_name)
         time_delta_sending = dt.now() - first_now_sending
         time_delta_counter = dt.now() - first_now_counter
         start_sending = time_delta_sending.total_seconds() > timeout_sending
         reset_counter = time_delta_counter.total_seconds() > timeout_counter
         if reset_counter:
            i += 1
            attachments_name.append(im_name)
            first_now_counter = dt.now()
            beep(1,2)

         if (start_sending and (i > n_attachments_per_mail)):
            first_now_sending = dt.now()
            print('sending email with attachments:')
            print(attachments_name)
            beep(2,1)
            thread = Thread(target=send_mail, args=(attachments_name,))
            thread.start()
            # async_result = pool.apply_async(send_mail, (attachments_name,))
            i=0
            attachments_name = []


   cam0.release()
   del (cam0);
