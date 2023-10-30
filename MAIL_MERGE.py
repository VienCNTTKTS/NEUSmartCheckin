import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

df = pd.read_excel("C:/Users/FPTSHOP/Desktop/DS_DangKiThamGiaChaoTanK65.xlsx")

sender_email = "btc.chaotansitde@gmail.com"

size = len(df.index)
count = 0
for i in range(1):
    record = df.iloc[i]
    receiver_email = record["Email"]

    message = MIMEMultipart("alternative")
    message["Subject"] = "XÁC NHẬN ĐĂNG KÝ THAM DỰ THÀNH CÔNG LỄ CHÀO TÂN SINH VIÊN KHÓA 65 VIỆN CNTT&KTS"
    message["From"] = sender_email
    message["To"] = receiver_email

    html = """\
<p style="text-align: justify;"><strong><em><span style="font-size:12pt;">Th&acirc;n gửi """ + str(record["Ten"]) + """,</span></em></strong></p>
<p style="text-align: justify;"><span style="font-size:12pt;">Lời đầu ti&ecirc;n, BTC Lễ ch&agrave;o t&acirc;n sinh vi&ecirc;n kh&oacute;a 65 Viện CNTT&amp;KTS xin gửi lời cảm ơn ch&acirc;n th&agrave;nh nhất đến bạn v&igrave; đ&atilde; d&agrave;nh thời gian quan t&acirc;m v&agrave; đăng k&yacute; tham dự chương tr&igrave;nh. BTC xin được tr&acirc;n trọng th&ocirc;ng b&aacute;o:</span></p>
<p style="text-align: center;"><strong><span style="color:#a64d79;background-color:#ffffff;font-size:13pt;">BẠN Đ&Atilde; ĐĂNG K&Yacute; THAM DỰ TH&Agrave;NH C&Ocirc;NG</span></strong></p>
<p style="text-align: center;"><strong><span style="color:#a64d79;background-color:#ffffff;font-size:13pt;">LỄ CH&Agrave;O T&Acirc;N SINH VI&Ecirc;N KH&Oacute;A 65 VIỆN CNTT&amp;KTS</span></strong></p>
<p style="text-align: justify;"><span style="font-size:12pt;">BTC xin gửi đến bạn c&aacute;c th&ocirc;ng tin quan trọng của buổi lễ như sau:</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">+ Thời gian diễn ra:&nbsp;</span><span style="background-color:#ffffff;font-size:12pt;">18h15, thứ s&aacute;u ng&agrave;y 27/10/2023</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">+ Địa điểm:&nbsp;</span><span style="background-color:#ffffff;font-size:12pt;">Hội trường A2, Trường Đại học Kinh tế Quốc d&acirc;n</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">+ Thời gian check-in:&nbsp;</span><span style="background-color:#ffffff;font-size:12pt;">17h20, thứ s&aacute;u ng&agrave;y 27/10/2023</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">+ Điểm đo&agrave;n t&iacute;ch lũy: 5 điểm</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">Để chương tr&igrave;nh diễn ra th&agrave;nh c&ocirc;ng v&agrave; đảm bảo quyền lợi của tất cả c&aacute;c bạn sinh vi&ecirc;n đến tham dự, dưới đ&acirc;y l&agrave; một số lưu &yacute;:</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">&bull; Thời gian check-in từ&nbsp;</span><span style="background-color:#ffffff;font-size:12pt;">17h20</span><span style="font-size:12pt;">, bạn vui l&ograve;ng tham gia đ&uacute;ng giờ để việc l&agrave;m thủ tục check-in.&nbsp;</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">&bull; Để nhận được điểm đo&agrave;n, c&aacute;c bạn sinh vi&ecirc;n tham gia sự kiện bắt buộc thực hiện check-in, check-out đầy đủ. M&atilde; check in v&agrave; check out của c&aacute;c bạn sinh vi&ecirc;n được gửi qua link sau:</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;"><img src=\"""" + ("https://raw.githubusercontent.com/VienCNTTKTS/NEUSmartCheckin/main/events/Ch%C3%A0o%20t%C3%A2n%20K65%20Vi%E1%BB%87n%20CNTT%26KTS/students/SV_" + str(record["MSV"]) + ".png") + """\"/></span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">&bull; M&atilde; check in v&agrave; check out chỉ được sử dụng một lần ứng với MSV bạn đ&atilde; đăng k&yacute; tham dự chương tr&igrave;nh n&ecirc;n bạn vui l&ograve;ng kh&ocirc;ng gửi m&atilde; check in, check out c&aacute; nh&acirc;n cho người kh&aacute;c.&nbsp;</span></p>
<p style="text-align: justify;"><span style="font-size:12pt;">Mọi thắc mắc vui l&ograve;ng li&ecirc;n hệ cho BTC qua email n&agrave;y hoặc inbox qua fanpage&nbsp;</span><strong><em><span style="color:#050505;font-size:12pt;">Li&ecirc;n chi đo&agrave;n Viện C&ocirc;ng nghệ th&ocirc;ng tin v&agrave; Kinh tế số - NEU:&nbsp;</span></em></strong><a href="https://www.facebook.com/cnttkts"><u><span style="color:#1155cc;font-size:12pt;">https://www.facebook.com/cnttkts</span></u></a></p>
<p style="text-align: justify;"><strong><em><span style="font-size:12pt;">Tr&acirc;n trọng!</span></em></strong></p>
<p style="text-align: justify;"><em><span style="font-size:13pt;">&nbsp;-------------------------------------------</span></em></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;">CHUỖI HOẠT ĐỘNG CH&Agrave;O T&Acirc;N SINH VI&Ecirc;N KH&Oacute;A 65: &ldquo;𝗪𝗮𝗻𝗱𝗲𝗿𝗹𝘂𝘀𝘁&rdquo;</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;"><span style="border:none;"><img alt="🔶" src="https://lh7-us.googleusercontent.com/9C8tcaQfKUBZk0IpB8bVCBFR6xg5_ag52koCDuCDGs_pRiHkVO4_fVYBjt19Bq3haCne24I2KPZKiYep9zOcYpHr5mlUH1h8eBKeveKIq-NFuDSUZQdXsnTpiu5TL-YLZXJ4yG0GmEEeK3WHNfxr258" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">Teambuilding: 8/10/2023</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;"><span style="border:none;"><img alt="🔶" src="https://lh7-us.googleusercontent.com/ChTdWC_N6Jt_aXv9AWMw8k6dzinIgUzmn7T_R9RBcPOS18_oARIg9YErlvNvuHJ2fPyy4w6rz9vas8RDTBRhlfZteqMQewknHVcu2b3rehrA9hpAT2fl6uAfIDUvu6Dk94nHzFsliLzaANts8-mryg4" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">Tuyển CTV BCH LCĐ: 02/10/2023 - 31/10/2023</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;"><span style="border:none;"><img alt="🔶" src="https://lh7-us.googleusercontent.com/hSo00IE4cjpWnxEJjODE1ihcL6OjtHfe3iqo2LbIodjxMMz0oj6sPMUVyPvANsaK7la-JKK70UduLihrU2Os9iQj33pTKsID_Z-i6TmTFnpGhF1iV770tgM7PAnJIuvtqxgbACKWpGS0FHDh2TA85Wc" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">Cuộc thi l&agrave;m video giới thiệu lớp Kh&oacute;a 65: 01/10/2023 - 15/10/2023</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;"><span style="border:none;"><img alt="🔶" src="https://lh7-us.googleusercontent.com/-kGjFfkmZsQtXCokaGfI9eegwRghG6blwz7ISiCknMuJVtez6dkZ7XB9He0UQrjDx-q-XMQv926igk8hmtnVQn4sOPBG8OjhAnnt786Rrsl_jqBl7vdbrFSIeIx0cKnyHgjVIUuQkHPmTmIZHsQQtHo" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">Lễ ch&agrave;o t&acirc;n sinh vi&ecirc;n v&agrave; cuộc thi văn nghệ: 27/10/2023</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;">--------------------------------------</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;">CHUỖI CHƯƠNG TR&Igrave;NH CH&Agrave;O T&Acirc;N SINH VI&Ecirc;N KH&Oacute;A 65: &ldquo;𝗪𝗮𝗻𝗱𝗲𝗿𝗹𝘂𝘀𝘁&rdquo;</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;"><span style="border:none;"><img alt="▪️" src="https://lh7-us.googleusercontent.com/Cc4W4aehAO-NhMuFuBYFQz7DJpwgw4O0NiLYXJ6tf2hlo0k4ISKELdnt02neJRhGhgmE1WME-V4K7Fp91dTF95iqtovE4tNG3eo-beIZhOxC0x4BbNqAUYz5PrL2MoELDg3wzc-KSvrhUkCxBNCxj1M" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">&nbsp;Website:</span><a href="https://sitde.neu.edu.vn/?fbclid=IwAR357IYiOilkLDzU-ooMguJ0M2J_dRLjWeS374csQsOgAUe_GSC83Dv_iuI"><span style="color:#050505;font-size:12pt;">&nbsp;</span><span style="color:#1155cc;font-size:12pt;">https://sitde.neu.edu.vn/</span></a></p>
<p style="text-align: justify;"><span style="color:#1155cc;font-size:12pt;"><span style="border:none;"><img alt="▪️" src="https://lh7-us.googleusercontent.com/C1F_TTmd0ApCda-93Nqb0-84ft6xH6ucxuMbODeM9_9PKDq0SZi5NzyvBIGZpO7h0PyJktgelqrpqqdnd6PXv1Rvw0Ld2gqMBAGCoSUEmWH1aG-Swq2bolBTl3jg3pqB4rok6vPL0Vdkh3gDuo1IShI" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">&nbsp;Email: lcdviencntt-kts@st.neu.edu.vn</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;"><span style="border:none;"><img alt="▪️" src="https://lh7-us.googleusercontent.com/1Sb6SQE-fCuhwFR0otBjgOppUpYl-T9ko01o5QSOUD9bqbTMAod4gWpREn0Mah1H4EAVRKjggc6q_SiQy82IgcJe0Yjn26sgqG7QTjlr1q7Hxc_NpFjo4n7echRaYpdYc8EaMJRcdo7vjwMMgYNFkOc" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">&nbsp;Tiktok:</span><a href="https://vt.tiktok.com/ZSJ3XJ8Ja/?fbclid=IwAR0zyHbkbbx6Ln2TOSrLtNQXo0GPlikSj2Xqdy4zH0o4-iQdO3bUV5Ii10A"><span style="color:#1155cc;font-size:12pt;">https://vt.tiktok.com/ZSJ3XJ8Ja/</span></a></p>
<p style="text-align: justify;"><span style="color:#1155cc;font-size:12pt;"><span style="border:none;"><img alt="▪️" src="https://lh7-us.googleusercontent.com/USp_gNj5IRImxy00VKiJlIzi4hbYTZvK5sUFt4QfM6WEvxQk1wCakqAjTj7WIR3b9ijz20UsRv6U7DIb0Dl9QR19exc_DbPd8LlAxLpV7pqMN1KM6J9xAzdMIcS5LTFqkr0143VcWseBtLlAE7vlvWw" width="16" height="16"></span></span><span style="color:#050505;font-size:12pt;">&nbsp;Hotline:</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;">Trưởng Ban tổ chức - Nguyễn Thu&yacute; Hiền (SĐT: 0916622388)</span></p>
<p style="text-align: justify;"><span style="color:#050505;font-size:12pt;">Ph&oacute; Ban tổ chức - Đinh Thu Phương (SĐT: 0399364008)</span></p>
<p><br></p>
<p><br></p>
    """

    part2 = MIMEText(html, "html")

    message.attach(part2)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.sendmail(
            sender_email, receiver_email, message.as_string()
        )
    print(str(i) + "/" + str(size) + " " + receiver_email)