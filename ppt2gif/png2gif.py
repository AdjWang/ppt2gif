# -*- coding: UTF8 -*-

import re
import imageio
from pathlib2 import Path

def get_png_num(file_path):
    """ This is a sort_key, sorting png files by file number
    png files are named as:
        幻灯片1.png
        幻灯片2.png
        ...
        幻灯片16.png
        幻灯片17.png
        ...
    These files are listed by str after glob('*.png') as:
        幻灯片1.png
        ...
        幻灯片16.png
        幻灯片17.png
        ...
        幻灯片2.png
        ...
    It will cause a wrong sequence of gif frames. 
    Use this function as a sort key to sort these files.
    """
    re_res = re.findall(r'.*?([\d]+).*?', Path(file_path).stem)
    if re_res and re_res[-1]:
        return int(re_res[-1])
    else:
        raise FileNotFoundError('png file number parse error')

def get_all_pngs(folder, sort_key=get_png_num):
    """ list all of png files in a folder """
    assert Path(folder).is_dir()
    return list(sorted(Path(folder).glob('*.png'), key=sort_key))

def png2gif(input_pngs, output_gif, duration=1, loop=0):
    """ convert pngs into gif
    
    Args:
        input_pngs: List of pngs to convert. Image will be read in sequence.
        output_gif: Output gif file path.
        duration: Duration of the frame. By second.
        loop: Number of loop. Infinite if loop = -1
    """
    images = [imageio.imread(str(file)) for file in input_pngs]
    imageio.mimsave(output_gif, images, duration=duration, loop=loop)

if __name__ == "__main__":
    pngs = get_all_pngs(r"D:\PythonProjects\ppt2gif\ppt2gif\ppts\L1")
    print(pngs)
    png2gif(pngs, r"D:\PythonProjects\ppt2gif\ppt2gif\ppts\L1.gif", 1, -1)
