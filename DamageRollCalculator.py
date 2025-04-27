from tkinter import ttk
import random

 
def build_damage_calc(self):
        container = ttk.Frame(self.damage_page, padding=10)
        container.pack(expand=True, fill='both')

        label = ttk.Label(container, text="Damage Calculator", font=("Arial", 16, "bold"))
        label.pack(pady=(0, 10))

        ttk.Label(container, text="Enter Number of Dice:").pack(pady=5)
        self.dice_count_entry = ttk.Entry(container, width=20)
        self.dice_count_entry.pack(pady=5)
        
        ttk.Label(container, text="Enter Sides of Die:").pack(pady=5)
        self.dice_sides_entry = ttk.Entry(container, width=20)
        self.dice_sides_entry.pack(pady=5)
        
        ttk.Label(container, text="Enter Modifier(+ , -):").pack(pady=5)
        self.modifier_entry = ttk.Entry(container, width=20)
        self.modifier_entry.pack(pady=5)

        calc_button = ttk.Button(container, text="Calculate Damage", command=self.calculate_damage)
        calc_button.pack(pady=10)

        self.result_label = ttk.Label(container, text="Result: ", font=("Arial", 12))
        self.result_label.pack(pady=20)

def calculate_damage(self):
        try:
            #get the values 
            dice_count = int(self.dice_count_entry.get())
            dice_sides = int(self.dice_sides_entry.get())
            modifier = int(self.modifier_entry.get()) if self.modifier_entry.get() else 0
            
            
            dice_rolls = [random.randint(1, dice_sides) for _ in range(dice_count)]
            total_roll = sum(dice_rolls)

            # calculate the total damage
            total_damage = total_roll + modifier

            #display 
            self.result_label.config(text=f"Result: {total_damage} (You Rolled: {', '.join(map(str, dice_rolls))}, Modifier: {modifier})")
    
        except ValueError:
            self.result_label.config(text="Error: Please enter valid integers for all fields.")