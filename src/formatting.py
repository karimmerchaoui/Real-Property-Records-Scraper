def format_house_feature(input_string, features):
    # Split the input string into lines
    lines = input_string.strip().split('\n')

    # Initialize an empty dictionary to hold the results
    result_dict = {}

    # Iterate through the list of words
    for word in features:
        # Iterate through the lines to find the word
        for i in range(len(lines) - 1):  # Ensure we don't go out of bounds
            if word in lines[i]:
                result_dict[word] = lines[i + 1].strip()  # Store the next line as value
                break  # Stop searching after finding the first occurrence

    return result_dict