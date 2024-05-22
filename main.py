import pandas as pd
from PIL import Image, ImageDraw, ImageFont

data = pd.read_excel (r'gen.xlsx')

# print(data['name1'])
name1_list = data["name1"].tolist()
name2_list = data["name2"].tolist()
patronymic_list = data["patronymic"].tolist()
status_list = data["status"].tolist()
kvant_list = data["kvantum"].tolist()
mod_list = data["mod"].tolist()
hour_list = data["hour"].tolist()
number_list = data["number"].tolist()





count = len(name1_list)
for i in range(len(name1_list)):
    sert_type = 0

    im_nice = Image.open(r'с отличием.png')
    im = Image.open(r'обычный.png')
    im_none = Image.open(r'Справка.png')
    width, height = im.size


    if status_list[i].lower() == 'сертификат с отличием':
        status_list[i] = "Успешно прошёл(-ла) обучение по\n" \
                         "дополнительной общеобразовательной\n" \
                         "общеразвивающей программе"
        sert_type = 0
        count -= 1
    elif status_list[i].lower() == 'сертификат':
        status_list[i] = "прошёл(-ла) обучение по\n" \
                         "дополнительной общеобразовательной\n" \
                         "общеразвивающей программе"
        sert_type = 1
        count -= 1
    elif status_list[i].lower() == 'справка':
        status_list[i] = "прослушал(-ла) курс по\n" \
                         "дополнительной общеобразовательной\n" \
                         "общеразвивающей программе"
        sert_type = 2
        count -= 1
    else:
        count -= 1
        continue

    txt = status_list[i].split()
    if sert_type == 0:
        d = ImageDraw.Draw(im_nice)
        text_color = (84, 150, 155)
    elif sert_type == 2:
        d = ImageDraw.Draw(im_none)
        text_color = (124, 153, 178)
    else:
        d = ImageDraw.Draw(im)
        text_color = (124, 153, 178)
    font = ImageFont.truetype("fonts/RussoOne-Regular.ttf", 70)

    name = f'{name1_list[i]} {name2_list[i]}'


    # ___________name______________
    _, _, w, h = d.textbbox((0, 0), name, font=font)
    d.text(((width - w) / 2 + 350, (400)), name.upper(), font=font, align ="center", fill=text_color)

    # ___________patr______________
    if type(patronymic_list[i]) == float:
        patronymic_list[i] = ''
    _, _, w, h = d.textbbox((0, 0), patronymic_list[i], font=font)
    d.text(((width - w) / 2 + 350, (470)), patronymic_list[i].upper(), font=font, align ="center", fill=text_color)

    # ___________status______________

    if type(status_list[i]) == float:
        status_list[i] = ''
    status = str(status_list[i])
    _, _, w, h = d.textbbox((0, 0), status, font=font)
    d.text(((width - w) / 2 + 750, (600)), status, font=ImageFont.truetype('fonts/Roboto-Light.ttf',40), align ="center", fill=(0, 0, 0))

    # ___________kvant______________
    _, _, w, h = d.textbbox((0, 0), kvant_list[i], font=font)
    d.text(((width - w) / 2 + 500, (770)), kvant_list[i], font=ImageFont.truetype('fonts/Roboto-Light.ttf',35), align ="center", fill=(0, 0, 0))

    # ___________mod______________
    if type(mod_list[i]) == float:
        mod_list[i] = 'базовый модуль'
    _, _, w, h = d.textbbox((0, 0), mod_list[i], font=font)
    d.text(((width - w) / 2 + 560, (730)), mod_list[i].lower(), font=ImageFont.truetype('fonts/Roboto-Light.ttf',35), fill=(0, 0, 0))

    # ___________hour______________
    hour = f'Объем часов - {hour_list[i]} ч.'
    _, _, w, h = d.textbbox((0, 0), hour, font=font)
    d.text(((width - w) / 2 + 100, (850)), hour, font=ImageFont.truetype('fonts/Roboto-Light.ttf',35), fill=(0, 0, 0))

    # ___________number______________
    if type(number_list[i]) == float:
        number_list[i] = ''
    _, _, w, h = d.textbbox((0, 0), number_list[i], font=font)
    d.text(((width - w) / 2 - 500, (1350)), number_list[i], font=ImageFont.truetype('fonts/Roboto-Light.ttf',25), fill=(20, 20, 20))

    if "успешно" in txt:
        im_nice.save("res/certificate_" + name + ".pdf")
    elif sert_type == 2:
        im_none.save("res/certificate_" + name + ".pdf")
    else:
        im.save("res/certificate_" + name + ".pdf")
    print(f"{count} + {name}")

