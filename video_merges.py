import os
import ffmpy3
from ffmpy3 import FFmpeg
from time import *

import codecs
def make_txt(path,filename):
    merge_path = path
    merge_filename = filename
    video_list = os.listdir(merge_path)
    for y in video_list:
        suffix = y.split('.')[-1]
        if suffix == 'txt' or suffix == 'm3u8':
            video_list.remove(y)

    sorted(video_list)
    txt_path = os.path.join(merge_path,merge_filename)
    if not os.path.exists(txt_path):
        f = codecs.open(txt_path,'w')
        for i in video_list:
            path1 = os.path.join(path,i)
            f.writelines(f'file \'{path1}\'')  # linux环境下需注意'\'与'/'
            f.writelines('\n')
        f.close()

def video_merge(txt_path,videopath):
    txt = txt_path
    input_file = txt
    output_file = videopath
    cmd = '-vcodec copy '
    ff = FFmpeg(inputs={input_file: '-f concat -safe 0 '}, outputs={output_file: cmd})
    if os.path.exists(output_file):
        pass
    else:
        try:
            ff.run()
        except:
            with open(r'E:\withai_document\视频管理\LC10000\视频合并\LPD视频合并\error.txt', "a+") as f:
                f.write(videopath  + " " + str(ff) + "\n")
            print('合并失败')
            print(ff)
        print(ff)

if __name__ == '__main__':
    ori_ = r'F:\video-translate\LC_videos\m3u8'
    save = r'F:\video-translate\LC_videos\mp4'
    video_name = '2.mp4'
    txt_name = '2.txt'
    txt_path = os.path.join(ori_, txt_name)
    make_txt(path=ori_, filename=txt_name)
    begin_time = time()
    video_merge(txt_path=txt_path, videopath=os.path.join(save, video_name))
    end_time = time()
    run_time = end_time - begin_time
    print(begin_time)
    print(end_time)
    print(run_time)

