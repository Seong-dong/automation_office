import os

if __name__ == "__main__":
    #root_dir = "//192.168.0.100/png/MD1/_/FANATICS/01. Main Order/2022 PO (106266~)"
    #root_dir = "//192.168.0.100/trade/민지/2022 수출서류작성건"
    root_dir = "//192.168.0.100/png/MD2/_/CHAMPION EUROPE/01. Main Order"
    for (root, dirs, files) in os.walk(root_dir):
        #print("# root : " + root)
        """
        if len(dirs) > 0:
            for dir_name in dirs:
                print("dir: " + dir_name)
        """
        percentChk = 1
        #if len(files) > 0:
        for file_name in files:
            if file_name.endswith(".xls"):
                #print(os.path.abspath(file_name))
                fileName = file_name #파일이름할당
                if ("F22PNGR1UK44" in fileName):
                    #print(os.path.abspath(file_name))
                    print(root + '/' + file_name)

print("finish")

#os.rename('//192.168.0.100/png/MD1/_/FANATICS/01. Main Order/2022 PO (106266~)/106269-FLC/106269-FLC REV042222.xlsx', '//192.168.0.100/png/MD1/_/FANATICS/01. Main Order/2022 PO (106266~)/106269-FLC/106269-FLC@REV04222022.xlsx')
