from copy import copy
def test(input):
    input = copy(input)
    input[0][1] = 10
    print(input)

hello = [[1,2,3],2,3]
test(hello)
#
print(hello)