def list_duplicates(seq):
    seen = set()
    seen_add = seen.add
    # adds all elements it doesn't know yet to seen and all other to seen_twice
    seen_twice = set(x for x in seq if x in seen or seen_add(x))
    # turn the set into a list (as requested)
    return list(seen_twice)

# Define a function which will replace my postfix string and return only the original text of servername
def replace_digits(text):
    import re
    return re.sub(r'dimofinf[0-9]*', '', text)