# -*- coding: utf-8 -*-
""" Convert ppt into gif 

Use PPT to make a form or diagram is effectively.
But sometimes PPT is hard to embed into HTML, while a gif maybe a better idea.
It is very useful if you need a gif to show difference in the change of forms, diagrams...

This tool can only running on Windows! Because the PPT handle package is win32com.

Usage:
    # The ppt_path can be a folder or a .pptx path or a list of .pptx path.
    # Such like:
    # "C:\\Users\\Administrator\\Desktop\\myPPTs"
    # "C:\\Users\\Administrator\\Desktop\\myPPTs\\temp.pptx"
    # ["C:\\Users\\Administrator\\Desktop\\myPPTs\\temp.pptx", ...]
    
    ppt_path = "C:\\Users\\Administrator\\Desktop\\myPPTs"
    import ppt2gif
    ppt_obj = ppt2gif.PPT(ppt_path)
    ppt_obj.convert2gif(duration=1, loop=-1)    # gif loop infinitely if loop=-1

The directory of the gif converted is the same as the .pptx.
Or you only need the .png images, use:
    ppt_obj.convert2png()
then the images will be put into a folder.

If you need both of the png and the gif, use:
    ppt_obj.convert2gif(duration=1, loop=-1, del_png=False)

help(ppt2gif.PPT) for more detail.
"""

from .ppt2png import PPT

__all__ = ['PPT']
