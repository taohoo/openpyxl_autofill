
def _duplicate(l):
    """
    list去重
    :param l:
    :return: 去重后的list
    """
    unique_list = []
    [unique_list.append(x) for x in l if x not in unique_list]
    return unique_list
