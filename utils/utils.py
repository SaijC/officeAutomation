class stringUtils:
    """
    class to handle string operations
    """

    def replaceSpace(self, inputStr):
        """
        Replace any space character with '_'
        :param inputStr: string
        :return: strin
        """
        if ' ' in inputStr:
            inputStr = inputStr.replace(' ', '_')
        return inputStr