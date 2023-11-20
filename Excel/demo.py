from PIL import Image
import xlsxwriter
Max_Width=500

def RGB_to_Hex(colors):
    item = colors.split(',')            # 将RGB格式划分开来
    color = '#'
    for i in item:
        num = int(i)
        # 将R、G、B分别转化为16进制拼接转换并大写  hex() 函数用于将10进制整数转换成16进制，以字符串形式表示
        color += str(hex(num))[-2:].replace('x', '0').upper()
    return color
def img_to_excel(pic_file,output_file='pic.xlsx',max_width=Max_Width):

    img = Image.open(pic_file)
    img_height=img.height
    img_width=img.width
    if img.width>max_width:
        # img_height=int(max_width/img.width*img.height)
        # img_width=max_width
        # img.resize((img_width,img_height))
        newwidth = int(max_width * img.width)
        newheight = int(max_width * img.height)
        img.resize((newwidth, newheight))


    workbook = xlsxwriter.Workbook(output_file)  #创建工作蒲
    worksheet = workbook.add_worksheet()  #创建表
    #设置行高和列宽
    worksheet.set_default_row(9)
    worksheet.set_column(0,img_width, 1)

    for i in range(img_width):
        for j in range(img_height):
            color=tuple([item - item%6 for item in img.getpixel((i, j))])
            rgb = str(color[0]) + ',' + str(color[1]) + ',' + str(color[2])

            cell_format = workbook.add_format({'bg_color':RGB_to_Hex(rgb) })
            worksheet.write(j, i, '', cell_format)
    workbook.close()


img_to_excel("./孙悟空.jpeg", 'pic.xlsx', 0.8)