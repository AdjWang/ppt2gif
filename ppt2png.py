# -*- coding: utf-8 -*-
"""	convert ppt in png or gif """

from tqdm import tqdm
from pathlib2 import Path
from contextlib import contextmanager
try:
    import win32com
    import win32com.client
except:
    raise ImportError('No module named win32com, use "pip install pypiwin32" to install.')

import sys
sys.path.append(str(Path(__file__).parent))
import png2gif

class PPT(object):
    """ ppt converter class
    
    Args:
        file_path: ppt path. A folder path, a .ppt or .pptx file path, 
        a list of ppt file path are both acceptable.
        
        visible: Display ppt window. Must be True, or a error will be raised.
    
    Attributes:
        ppts: list of ppt file pathes.
        file_count: count of ppt files.
        
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
    
    References:
        https://www.jianshu.com/p/6b99ee1b845e
    
    """
    def __init__(self, file_path, visible=True):
        # check the type of file_path
        if isinstance(file_path, list):                  # list
            self.ppts = list(map(Path, file_path))
        else:                                            # str
            if Path(file_path).is_dir():                     # dir of file_path
                self.ppts = self.get_all_ppts(file_path)
            else:                                            # single file
                self.ppts = [Path(file_path)]
        
        # check if ppt file was found   
        self.file_count = len(self.ppts)
        if self.file_count == 0:
            raise FileNotFoundError('ppt file not found!')
        
        self.__powerpoint = win32com.client.Dispatch('PowerPoint.Application')
        self.__powerpoint.Visible = visible
    
    @staticmethod
    def get_all_ppts(folder):
        """ list all of ppt files in a folder """
        assert Path(folder).is_dir()
        return list(Path(folder).glob('*.ppt')) + list(Path(folder).glob('*.pptx'))

    def __str__(self):
        """ print list of ppt files and count number
        
        usage:  
            obj = PPT(PPT_FILE_PATH)
            print(obj)
        """
        return '\n'.join(list(map(str, self.ppts))) \
                 + f'\nTotal: {self.file_count}'
    
    @contextmanager
    def open(self, file_name):
        """ open a ppt file
        Usage:
            with self.open(file_name) as f:
                ...
                
        Args:
            file_name: A ppt file path, end with .ppt or .pptx.
        
        Raises:
            FileNotFoundError: Raised if the ppt file is not found.
        """
        try:
            self.ppt = self.__powerpoint.Presentations.Open(file_name)
            yield self.ppt
            self.ppt.Close()
        except:
            raise FileNotFoundError('ppt file not found!')
        
    def convert2png(self):
        """ convert ppt into png 
        
        ppt files have been listed in self.ppts when instancing.
        The process to convert a ppt named temp.pptx:
            1. open the temp.pptx with function 'self.open'
            2. make a directory named temp
            3. save pngs into the directory
            4. close temp.pptx
            5. return [temp]
        
        Converted ppt will be ignored, you can continue to convert at any time.
        
        Returns:
            png_folders: List of directories made. With the same name of ppt.
            
        """
        # tried to use concurrent package but failed for FileNotFoundError raised...
        png_folders = []
        for path in tqdm(self.ppts, desc='ppt to png'):
            output_folder = path.parent.joinpath(path.stem)
            png_folders.append(output_folder)   # recoder of output folders
            if output_folder.exists():
                continue
            with self.open(path) as f:          # open ppt and convert
                f.SaveAs(str(output_folder) + '.png', 18)
        return png_folders
    
    def convert2gif(self, duration, loop=0, del_png=True):
        """ convert ppt into gif
        
        Convert ppt into png first, then convert png into gif.
        
        Args:
            duration: Duration of the frame. By second.
            loop: Number of loop. Infinite if loop = -1
            del_png: Delete the folder of png if True
        
        """
        result_folders = self.convert2png()
        for folder in tqdm(result_folders, desc='png to gif'):
            pngs = png2gif.get_all_pngs(folder)     # get pngs from folder
            gif = Path(str(folder) + '.gif')        # make gif path
            if gif.exists():
                continue
            png2gif.png2gif(pngs, str(gif), duration, loop)     # save as gif
            if del_png:
                for png in pngs:
                    png.unlink()
                folder.rmdir()
    
    def __del__(self):
        self.__powerpoint.Quit()
    

if __name__ == "__main__":
    # ppts = PPT.get_all_ppts(r"D:\PythonProjects\ppt2gif\ppt2gif\ppts")
    # ppt = PPT(ppts)
    # ppt.convert2png()
    ppt = PPT(r"D:\PythonProjects\ppt2gif\ppt2gif\ppts")
    # ppt.convert2png()
    ppt.convert2gif(1, -1)
    # print(result_folders)
    
    