from typing import List


class SimpleDataUtils:
    @staticmethod
    def find_longest_list_length(list_of_lists: List[List]) -> int:
        longest = len(list_of_lists[0])
        for lst in list_of_lists:
            if len(lst) > longest:
                longest = len(lst)

        return longest
