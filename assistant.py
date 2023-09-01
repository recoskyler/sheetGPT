#!/usr/bin/python

import openai
from re import sub, search

__all__ = ["get_formatted_prompt", "get_answer"]

def get_formatted_prompt(prompt: str, input_values: object) -> str:
    formatted_prompt = prompt

    while search("(\$)([0-9]+)", formatted_prompt) != None:
        match = search("(\$)([0-9]+)", formatted_prompt).group()
        index = int(match[1:])

        if index < len(input_values):
            formatted_prompt = sub("(\$)([0-9]+)", input_values[index], formatted_prompt, 1)
        else:
            print("Found invalid placeholder, replacing with empty string: " + match)

            formatted_prompt = sub("(\$)([0-9]+)", "", formatted_prompt, 1)

            break

    return formatted_prompt

def get_answer(prompt: str, system_prompt: str, model: str, api_key: str) -> tuple:
    openai.api_key = api_key

    try:
        print("Getting answer from ChatGPT...")

        res = openai.ChatCompletion.create(model=model, messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": prompt}]).choices[0].message.content
    except Exception as e:
        print("Failed to get answer: ")
        print(e)
        return ("", 0)

    return (res, 1)
