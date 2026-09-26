import json
import os

def load_user_data(user_id):
    filename = user_id + ".json"
    if not os.path.exists(filename):
        return {}
    with open(filename, "r") as f:
        data = json.load(f)
    return data

def calculate_stats(numbers):
    total = 0
    count = 0
    for n in numbers:
        total += n
        count += 1
    avg = total / count
    max_val = 0
    for n in numbers:
        if n > max_val:
            max_val = n
    return {"avg": avg, "max": max_val, "count": count}

def format_report(stats):
    output = "=== Report ===\n"
    output += f"Average: {stats['avg']}\n"
    output += f"Max: {stats['max']}\n"
    output += f"Count: {stats['count']}\n"
    return output

if __name__ == "__main__":
    numbers = [10, 20, 30, 40, 50]
    stats = calculate_stats(numbers)
    print(format_report(stats))
