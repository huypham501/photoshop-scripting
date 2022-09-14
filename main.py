import os
import win32com.client

FILE_FORMAT_TYPE = (".psd", ".PSD",
                    ".psdc", ".PSDC",
                    ".png", ".PNG",
                    ".jpeg", ".JPEG",
                    ".tiff", ".TIFF",
                    ".heic", ".HEIC",
                    ".psb", ".PSB",
                    ".raf", ".RAF",
                    ".raw", ".RAW")



def is_exists_path(path: str):
    return os.path.exists(path)


def is_support_format_path(path: str):
    return path.endswith(FILE_FORMAT_TYPE)


class PhotoshopDocument:
    def __init__(self, photoshop_app: win32com.client.CDispatch, doc_path: str = None):
        self.photoshop_app = photoshop_app
        self.doc_path = None
        self.set_doc_path(doc_path)
        self.is_open = False
        self.ref = None

    def set_doc_path(self, doc_path: str):
        if not is_exists_path(doc_path):
            # Raise error
            print("not is_exists_path")
            return
        if not is_support_format_path(doc_path):
            # Raise error
            print("not is_support_format_path")
            return
        self.doc_path = doc_path

    def open(self):
        self.ref = self.photoshop_app.Open(self.doc_path)
        self.is_open = True

    def save_as_PNG(self, save_path: str, compression_ratio: int, interlaced: bool, as_copy: bool,
                    filename_low_case: int):
        if filename_low_case != 2 and filename_low_case != 3:
            # Raise error
            print("filename_low_case not support")
            return
        save_options = win32com.client.Dispatch("Photoshop.PNGSaveOptions")
        save_options.Compression = compression_ratio
        save_options.Interlaced = interlaced
        self.ref.SaveAs(save_path, save_options, as_copy, filename_low_case)

    def close(self, save_option: int):
        if 1 > save_option > 3:
            # Raise error
            print("save_option not support")
            return
        self.ref.Close(save_option)


def traveler_convert_to_PNG(ps_app: win32com.client.CDispatch, start_dir: str, from_format: str,
                            include_all_sub_folder: bool):
    print("Start session:" + start_dir)
    if not is_exists_path(start_dir):
        # Raise error
        return

    ps_app.Purge(4)

    for root, dirs, file_names in os.walk(start_dir):
        print("Processing in " + start_dir)
        if len(file_names) == 0:
            return
        for index, file_name in enumerate(file_names):
            print(str(index + 1) + "/" + str(len(file_names)))
            if file_name.endswith(from_format):
                print("Convert")
                doc = PhotoshopDocument(ps_app, root + '\\' + file_name)
                doc.open()
                doc.save_as_PNG(root + '\\' + file_name, 5, True, True, 2)
                doc.close(2)
                ps_app.Purge(4)
                continue
            print("Skip")
        if not include_all_sub_folder:
            break
    print("Complete session")
    print("--------------------------------------------------------------")

if __name__ == '__main__':
    ps_app = win32com.client.Dispatch("Photoshop.Application")
    traveler_convert_to_PNG(r"D:\Example", ".RAF", True)
