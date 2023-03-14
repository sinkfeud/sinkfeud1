# 导入所需的模块
import poplib
import email
import os

# 定义一个函数，用于解析邮件内容，并返回发件人地址、收件人地址、邮件主题、邮件正文和附件列表
def parse_email(msg):
    # 获取发件人地址
    from_addr = email.utils.parseaddr(msg.get('From'))[1]
    # 获取收件人地址
    to_addr = email.utils.parseaddr(msg.get('To'))[1]
    # 获取邮件主题
    subject = msg.get('Subject')
    # 解码邮件主题，如果有编码信息的话
    subject = email.header.decode_header(subject)[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()
    # 初始化邮件正文和附件列表为空字符串和空列表
    body = ''
    attachments = []
    # 遍历邮件的所有部分（可能有多个）
    for part in msg.walk():
        # 判断部分类型是否为文本或HTML
        if part.get_content_type() == 'text/plain' or part.get_content_type() == 'textml':
            # 获取部分内容，并解码为字符串，如果有编码信息的话
            content = part.get_payload(decode=True)
            charset = part.get_content_charset()
            if charset:
                content = content.decode(charset, errors='ignore')
            # 拼接到邮件正文中
            body += content + '\n'
        # 判断部分类型是否为附件（有文件名）
        elif part.get_filename():
            # 获取文件名，并解码为字符串，如果有编码信息的话
            filename = part.get_filename()
            filename = email.header.decode_header(filename)[0][0]
            if isinstance(filename, bytes):
                filename = filename.decode(errors='ignore')
            filepath = os.path.join('C:\\Users\\Bara.san\\Desktop\\python', filename)
            # 打开文件，并将部分内容写入文件中，关闭文件
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            # 将文件路径添加到附件列表中
            attachments.append(filepath)
    
    return from_addr, to_addr, subject, body, attachments

# 定义一个函数，用于搜索邮箱中包含指定关键词的邮件，并打印相关信息和下载附件（如果有）
def search_email(keyword):
    # 连接到POP3服务器（这里以gmail为例），并输入用户名和密码进行登录（请替换成你自己的账户信息）
    pop_conn = poplib.POP3_SSL('pop.qq.com')
    pop_conn.user('3056181950@qq.com')
    pop_conn.pass_('tmavcywhetgydfbc')
    
    try:
        # 获取服务器上所有邮件的编号列表（从1开始）
        num_list = [int(i.split()[0]) for i in pop_conn.list()[1]]
        
        for num in num_list:
            
             try:
                 print(f'正在处理第{num}封邮件...')
                 print('-'*50)
                 # 从服务器上获取第num封邮件的原始数据（字节串）
                 raw_data = b'\n'.join(pop_conn.retr(num)[1])
                 # 将原始数据解析为Message对象（类似于字典）
                 msg = email.message_from_bytes(raw_data)
                 # 调用之前定义的函数，解析Message对象，并返回相关信息和附件列表 
                 from_addr, to_addr, subject, body, attachments = parse_email(msg)
                 
                 if keyword.lower() in subject.lower():
                     print(f'发现匹配关键词"{keyword}"的邮件！')
                     print(f'发件人地址：{from_addr}')
                     print(f'收件人地址：{to_addr}')
                     print(f'邮件主题：{subject}')
                     print(f'邮件正文：\n{body}')
                     if attachments:
                         print(f'邮件附件：')
                         for filepath in attachments:
                             print(filepath)
                     else:
                         print(f'邮件无附件。')
                 else:
                     print(f'未发现匹配关键词"{keyword}"的邮件。')
                 # 在每个邮件的输出结果之间用空白行做分隔符
                 print('\n')
             except Exception as e:
                 # 如果处理某封邮件时出现异常，打印异常信息，并继续处理下一封邮件
                 print(f'处理第{num}封邮件时出现异常：{e}')
                 continue
        
    finally:
        # 退出连接，释放资源
        pop_conn.quit()

# 定义一个函数，用于接收用户输入，并调用之前定义的函数进行搜索和输出
def main():
    while True:
        # 接收用户输入，并去除首尾空白字符
        keyword = input('请输入要搜索的关键词（输入"quit"或"exit"退出）：').strip()
        # 如果用户输入为空字符或空白字符，则继续提示
        if not keyword:
            continue
        # 如果用户输入为"quit"或"exit"（大小写不敏感），则退出程序
        elif keyword.lower() in ['quit', 'exit']:
            break
        else:
            # 调用之前定义的函数，进行搜索和输出
            search_email(keyword)

# 运行主函数
if __name__ == '__main__':
    main()