delivered = ['MS Question Randomization medium1', 'MS Question Randomization high1',
             'MS Question Randomization medium2', 'MS Question Randomization high5', 'MS Question Randomization Low 3',
             'MS Question Randomization Low 2']

expected_questions = ['MS Question Randomization Low 1', 'MS Question Randomization Low 2',
                      'MS Question Randomization Low 3', 'MS Question Randomization Low 4',
                      'MS Question Randomization Low 5',
                      'MS Question Randomization medium1', 'MS Question Randomization medium2',
                      'MS Question Randomization medium3', 'MS Question Randomization medium4',
                      'MS Question Randomization medium5', 'MS Question Randomization high1',
                      'MS Question Randomization high2', 'MS Question Randomization high3',
                      'MS Question Randomization high4', 'MS Question Randomization high5']

final = (set(delivered) - set(expected_questions))
print(len(final))
print(str(final))
#     print('got the question which is expected')
# else:
#     print('got the question which is not expected')
