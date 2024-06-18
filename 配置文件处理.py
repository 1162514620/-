import configparser

# 对配置文件的处理的类
class 配置文件处理:
    def __init__(self, filepath):
        self._path = filepath
        self.config = configparser.ConfigParser()  # 实例化解析对象
        self.config.read(filepath)  # 读文件

    def get_sections(self):
        """
        获取ini文件所有的块，返回为list
        """
        sect = self.config.sections()
        return sect

    def get_options(self, sec):
        """
        获取ini文件指定块的项
        :param sec: 传入的块名
        :return: 返回指定块的项（列表形式）
        """
        return self.config.options(sec)

    def get_items(self, sec):
        """
        获取指定section的所有键值对
        :param sec: 传入的块名
        :return: section的所有键值对（元组形式）
        """
        return self.config.items(sec)

    def get_option(self, sec, opt):
        """
        :param sec: 传入的块名
        :param opt: 传入项
        :return: 返回项的值(string类型)
        """
        return self.config.get(sec, opt)

    def write_(self):
        """ 将修改后写入文件 """
        with open(self._path, 'w') as fp:
            self.config.write(fp)

    def add_section(self, sec):
        """
        为ini文件添加新的section, 如果section 已经存在则抛出异常
        :param sec: 传入的块名
        :return: None
        """
        self.config.add_section(sec)
        self.write_()

    def set_option(self, sec, opt, value):
        """
        对指定section下的某个option赋值
        :param sec: 传入的块名
        :param opt: 传入的项名
        :param value: 传入的值
        :return:  None
        """
        self.config.set(sec, opt, value)
        self.write_()  # 写入文件

    def remove_sec(self, sec):
        """
        删除某个section
        :param sec: 传入的块名
        :return: bool
        """
        self.config.remove_section(sec)
        self.write_()  # 写入文件

    def remove_opt(self, sec, opt):
        """
        删除某个section下的某个option
        :param sec: 传入的块名
        :param opt: 传入的项名
        :return: bool
        """
        self.config.remove_option(sec, opt)
        self.write_()  # 写入文件
