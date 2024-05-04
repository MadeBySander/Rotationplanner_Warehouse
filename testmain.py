
# Static stations display with colors
for i, station in enumerate(stations):
    if i < len(stations) // 2:
        bg_color = 'lightyellow' if i % 2 == 0 else 'lightcyan'  # Alternating background colors
        tk.Label(window, text=station, bg=bg_color).grid(row=i + 0, column=3, padx=pad_x, pady=pad_y)

        # Add label for workers under each station
        workers_label = tk.Label(window, text=", ".join(workers['Day Shift'].keys()), bg=bg_color, wraplength=200)
        workers_label.grid(row=i + 0, column=4, padx=pad_x, pady=pad_y)
    else:
        j = i - len(stations) // 2
        bg_color = 'lightyellow' if j % 2 == 0 else 'lightcyan'  # Alternating background colors
        tk.Label(window, text=station, bg=bg_color).grid(row=j + 0, column=5, padx=pad_x, pady=pad_y)

        # Add label for workers under each station
        workers_label = tk.Label(window, text=", ".join(workers['Evening Shift'].keys()), bg=bg_color, wraplength=200)
        workers_label.grid(row=j + 0, column=6, padx=pad_x, pady=pad_y)


