
def split_invoice():
    # Ask for the total invoice amount
    try:
        total = float(input("Enter the invoice total amount: $"))
    except ValueError:
        print("Please enter a valid number.")
        return
    
    # Define the percentages for each property
    percentages = {
        "105": 47.45,
        "106": 14.9,
        "107": 37.65
    }
    
    # Calculate the amount for each property and round to 2 decimal places
    amounts = {}
    calculated_total = 0
    
    for prop, percentage in percentages.items():
        amount = round(total * percentage / 100, 2)
        amounts[prop] = amount
        calculated_total += amount
    
    # Check if there's a rounding error and adjust if needed
    difference = round(total - calculated_total, 2)
    if difference != 0:
        print(f"Rounding adjustment of ${difference} applied to property 107")
        amounts["107"] += difference
    
    # Display the results
    print("\nSplit invoice amounts:")
    print("---------------------")
    for prop, amount in sorted(amounts.items()):
        print(f"Property {prop}: ${amount:.2f}")
    print(f"Total: ${sum(amounts.values()):.2f}")

if __name__ == "__main__":
    print("Invoice Splitting Tool")
    print("=====================")
    split_invoice()
