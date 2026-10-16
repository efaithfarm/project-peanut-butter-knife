# Scrimba Python Intro Class #
# Project - Math Tutor

import random, time

time_scores = []
probs = {}
num_probs = int(input('Hi there! How many problems would you like to try? '))
high_end = int(input('Sure, no problem. What should we use as the highest number in the range?'))
num_range = range(1,high_end,1)
def tutor():
    prob_no = 1
    right = 0
    wrong = 0
    while prob_no <= num_probs:
        start = time.time()
        num1 = random.choice(num_range)
        num2 = random.choice(num_range)
        
        problem = f'{num1} x {num2}'
        
        answer = num1 * num2
        guess = int(input(f'What is {problem}? '))
        end = time.time()
        elapsed = round((end - start),1)
        if guess != answer:
            print(f'Sorry, that is incorrect. It took you {elapsed} seconds to come up with that awful answer.')
            wrong += 1
            probs[problem] = f'{str(guess)}, should be {answer}'
        else:
            print(f'Good job! You got it! Just {elapsed} seconds for that one!')
            right += 1
            probs[problem] = guess
        prob_no += 1
        time_scores.append(elapsed)
    score = int((right / num_probs) * 100)
    total_time = sum(time_scores)
    print(f'Your score: {right}/{num_probs} - {score} %! Overall time spent: {round(total_time,1)} seconds.')
    report = ''
    for prob in probs:
        report += '\n' + prob + ' = ' + str(probs[prob])
    print(f'Here is the report on your practice: {report}')   
tutor()
