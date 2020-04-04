# -*- coding: UTF8 -*-

import imageio
from pathlib2 import Path

def get_all_pngs(folder, sort_key=lambda p: int(p.stem[3:])):
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
    
