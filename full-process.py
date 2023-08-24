import os
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
import originpro as op
import pandas as pd
import sys


# Code snippet 1: Rename columns
def rename_columns():
    def rename_columns_action():
        # Select the CSV file
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])

        if file_path:
            try:
                # Prompt for start and end voltage values
                start_voltage = float(simpledialog.askstring("Start Voltage", "Enter the starting voltage:"))
                mid_voltage = float(simpledialog.askstring("Mid Voltage", "Enter the ending voltage:"))

                # Read the CSV file
                df = pd.read_csv(file_path)

                # Get the headers
                headers = df.columns.tolist()

                # Calculate the voltage step
                step_up = (mid_voltage - start_voltage) / (len(headers) // 2 - 1)
                step_down = (start_voltage - mid_voltage) / (len(headers) // 2 - 1)

                # Rename the columns
                for i in range(1, len(headers)):
                    if i <= len(headers) // 2:
                        voltage = start_voltage + (i - 1) * step_up
                    else:
                        voltage = mid_voltage + (i - len(headers) // 2 - 1) * step_down
                    new_header = "{:.2f}".format(voltage)  # Format the voltage with 2 decimal places
                    df.rename(columns={headers[i]: new_header}, inplace=True)

                # Save the renamed data to XLSX file
                save_path = os.path.splitext(file_path)[0] + "_renamed.xlsx"
                df.to_excel(save_path, index=False)

                # Open the folder where the output files are saved
                os.startfile(save_path)

                messagebox.showinfo("Success", "Columns renamed successfully. Saved as " + save_path)
            except Exception as e:
                messagebox.showerror("Error", "An error occurred: " + str(e))

    rename_columns_action()


root = tk.Tk()
root.withdraw()


# End of Code snippet 1: Rename columns

# code 2 start
def bg_processing():
    # Select the CSV file
    xlsx_file_path = filedialog.askopenfilename(title="Select Input XLSX File", filetypes=[("XLSX Files", "*.xlsx")])

    if xlsx_file_path:
        try:
            # Read the input XLSX file and rename the sheets
            xlsx = pd.ExcelFile(xlsx_file_path)
            sheets = xlsx.sheet_names
            df = None  # Initialize df variable
            wavenumber_column = None  # Initialize wavenumber_column variable

            for i, sheet in enumerate(sheets):
                df_sheet = pd.read_excel(xlsx_file_path, sheet_name=sheet)
                df_sheet.to_excel(xlsx_file_path, sheet_name=f"Sheet{i + 1}", index=False)

                if i == 0:
                    df = df_sheet
                    wavenumber_column = df_sheet["Wavenumber"]

            # Create a Tkinter window for column selection
            column_window = tk.Tk()
            column_window.title("Select Column")
            column_window.geometry("300x100")

            # Create a label for column selection
            column_label = ttk.Label(column_window, text="Choose a column:")
            column_label.pack()

            # Create a combobox for column selection
            column_combobox = ttk.Combobox(column_window, values=df.columns.tolist())
            column_combobox.pack()

            # Create a button to confirm column selection
            confirm_button = ttk.Button(column_window, text="Confirm", command=column_window.quit)
            confirm_button.pack()

            # Run the column selection window
            column_window.mainloop()
            # Get the chosen column
            chosen_column = column_combobox.get()

            # Create a new sheet for processing
            processed_sheet = pd.DataFrame()
            processed_sheet["Wavenumber"] = wavenumber_column

            for column in df.columns[1:]:
                if column == chosen_column:
                    processed_sheet[column] = 0
                else:
                    processed_sheet[column] = df[column] - df[chosen_column]

            # Prompt for the output directory
            xlsx_output_dir = filedialog.askdirectory(title="Select Output Directory")

            # Get the input filename without extension
            input_xlsx_file_name = os.path.splitext(os.path.basename(xlsx_file_path))[0]

            # Construct the output file path
            output_xlsx_file_name = f"{input_xlsx_file_name}_{chosen_column}.xlsx"
            output_xlsx_file_path = os.path.join(xlsx_output_dir, output_xlsx_file_name)

            # Save the processed sheet to a new workbook
            with pd.ExcelWriter(output_xlsx_file_path, engine="openpyxl") as writer:
                processed_sheet.to_excel(writer, sheet_name="Sheet1", index=False)

            # Open the folder where the output files are saved
            os.startfile(xlsx_output_dir)

            # Show completion message
            message = f"Processing completed! Output saved as:\n{output_xlsx_file_path}"
            messagebox.showinfo("Processing Complete", message)
            column_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", "An error occurred: " + str(e))


root = tk.Tk()
root.withdraw()


# Code snippet 2 end

# Code snippet 3: Create Origin graphs
def create_origin_graphs():
    def origin_shutdown_exception_hook(exctype, value, traceback):
        """Ensures Origin gets shut down if an uncaught exception"""
        op.exit()
        sys.__excepthook__(exctype, value, traceback)

    sys.excepthook = origin_shutdown_exception_hook

    # Only run if external Python
    if op.oext:
        # Create a new Origin project
        op.new()
        op.set_show(True)

    try:
        # Prompt user to select an Origin template file for graph 1
        template_file_path_1 = filedialog.askopenfilename(title="Select Origin Template File for Graph 1",
                                                          filetypes=[("Origin Template Files", "*.otpu")])

        # Prompt user to select a data file (CSV or XLSX) for the selected template
        data_file_types = [("CSV Files", "*.csv"), ("XLSX Files", "*.xlsx")]
        csv_file_path_1 = filedialog.askopenfilename(
            title=f"Select Data File for {os.path.basename(template_file_path_1)}",
            filetypes=data_file_types)

        # Load the CSV file into Origin worksheet
        wks = op.new_sheet()
        wks.from_file(csv_file_path_1, False)

        # Create a new graph using the selected template for graph 1
        gr1 = op.new_graph(template=template_file_path_1)
        gr1[0].add_plot(wks, 1, 0)
        gr1[0].rescale()

        # Add more graphs
        more_graphs = True
        graph_num = 2

        while more_graphs:
            # Prompt user to select an Origin template file for the graph
            template_file_path = filedialog.askopenfilename(title=f"Select Origin Template File for Graph {graph_num}",
                                                            filetypes=[("Origin Template Files", "*.otpu")])

            # Prompt user to select an XLSX file for the graph
            xlsx_file_path = filedialog.askopenfilename(
                title=f"Select XLSX File for {os.path.basename(template_file_path)}",
                filetypes=[("XLSX Files", "*.xlsx")])

            # Load the XLSX file into a new worksheet
            wks = op.new_sheet()
            wks.from_file(xlsx_file_path, False)

            # Create a graph page with the user-selected template for the graph
            gr = op.new_graph(template=template_file_path)
            gl = gr[0]

            # Prompt user for the range of columns to plot for the graph
            column_range = simpledialog.askstring("Range of Columns",
                                                  f"Enter the range of columns (e.g., 1-12) for Graph {graph_num}:")

            start_col, end_col = map(int, column_range.split("-"))
            y_columns = list(range(start_col, end_col + 1))

            for col in y_columns:
                plot = gl.add_plot(wks, col, 0)

            # Group and Rescale the graph
            gl.group()
            gl.rescale()

            # Prompt user if they want to add more graphs
            response = tk.messagebox.askyesno("Add More Graphs", "Do you want to add more graphs?")

            if response:
                graph_num += 1
            else:
                more_graphs = False

        # Tile all windows
        op.lt_exec('win-s T')

        # Prompt user for the output Origin file name
        output_file_path = filedialog.asksaveasfilename(title="Save Output Origin File",
                                                        filetypes=[("Origin Project Files", "*.opju")])

        # Creating a notes window
        nt = op.new_notes()

        # Appending input information to the notes
        nt.append("Input Information:")
        nt.append(f"Graph 1 Template: {os.path.basename(template_file_path_1)}")
        nt.append(f"Graph 1 CSV File: {os.path.basename(csv_file_path_1)}")

        for i in range(2, graph_num + 1):
            template_var = locals().get(f"template_file_path_{i}")
            csv_var = locals().get(f"csv_file_path_{i}")

            if template_var and csv_var:
                nt.append(f"\nGraph {i} Template: {os.path.basename(template_var)}")
                nt.append(f"Graph {i} CSV File: {os.path.basename(csv_var)}")

        # Appending folder path information to the notes
        nt.append("\nFolder Paths:")
        nt.append(f"Graph 1 Folder: {os.path.dirname(template_file_path_1)}")

        for i in range(2, graph_num + 1):
            template_var = locals().get(f"template_file_path_{i}")

            if template_var:
                nt.append(f"Graph {i} Folder: {os.path.dirname(template_var)}")

        nt.append(f"Output Folder: {os.path.dirname(output_file_path)}")

        # Appending output file information to the notes
        nt.append("\nOutput Information:")
        nt.append(f"Output File: {os.path.basename(output_file_path)}")

        # Displaying the note
        nt.view = 1

        # Tile all windows
        op.lt_exec('win-s T')

        # Save the project to the specified output file path
        if op.oext:
            output_file_path = os.path.abspath(output_file_path)
            op.save(output_file_path)

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        op.exit()


# End of Code snippet 3: Create Origin graphs


# Code snippet 4: Add graphs to existing Origin project
def add_graphs_to_project():
    def origin_shutdown_exception_hook(exctype, value, traceback):
        """Ensures Origin gets shut down if an uncaught exception"""
        op.exit()
        sys.__excepthook__(exctype, value, traceback)

    sys.excepthook = origin_shutdown_exception_hook

    # Only run if external Python
    if op.oext:
        op.set_show(True)

    def save_origin_project(output_path):
        try:
            op.save(output_path)
            return True  # Save operation successful
        except PermissionError:
            return False  # Save operation failed due to read-only
        except Exception as e:
            print(f"An error occurred while saving: {str(e)}")
            return False  # Save operation failed for other reasons

    def add_graphs_to_project_action():
        try:
            # Prompt user to select an existing Origin project file
            origin_file_path = filedialog.askopenfilename(title="Select Existing Origin Project File",
                                                          filetypes=[("Origin Project Files", "*.opju")])

            # Load the existing Origin project file
            op.open(origin_file_path)

            # Add more graphs
            more_graphs = True
            graph_num = 1

            while more_graphs:
                # Prompt user to select an Origin template file for the graph
                template_file_path = filedialog.askopenfilename(
                    title=f"Select Origin Template File for Graph {graph_num}",
                    filetypes=[("Origin Template Files", "*.otpu")])

                # Prompt user to select an XLSX file for the graph
                data_file_types = [("CSV Files", "*.csv"), ("XLSX Files", "*.xlsx")]
                xlsx_file_path = filedialog.askopenfilename(
                    title=f"Select XLSX File for {os.path.basename(template_file_path)}",
                    filetypes=data_file_types)

                # Load the XLSX file into a new worksheet
                wks = op.new_sheet()
                wks.from_file(xlsx_file_path, False)

                # Create a graph page with the user-selected template for the graph
                gr = op.new_graph(template=template_file_path)
                gl = gr[0]

                # Prompt user for the range of columns to plot for the graph
                column_range = simpledialog.askstring("Range of Columns",
                                                      f"Enter the range of columns (e.g., 1-12) for Graph {graph_num}:")

                start_col, end_col = map(int, column_range.split("-"))
                y_columns = list(range(start_col, end_col + 1))

                for col in y_columns:
                    gl.add_plot(wks, col, 0)

                # Group and Rescale the graph
                gl.group()
                gl.rescale()

                # Prompt user if they want to add more graphs
                response = tk.messagebox.askyesno("Add More Graphs", "Do you want to add more graphs?")

                if response:
                    graph_num += 1
                else:
                    more_graphs = False

            # Tile all windows
            op.lt_exec('win-s T')

            # Prompt user for the output Origin project file name and location
            output_file_path = filedialog.asksaveasfilename(title="Save Output Origin File As",
                                                            filetypes=[("Origin Project Files", "*.opju")])

            if not output_file_path:
                print("Save operation canceled by user.")
                op.exit()
                return  # Exit the function if the user cancels the save operation

            # Try to save the project file
            output_file_path = os.path.abspath(output_file_path)
            success = save_origin_project(output_file_path)

            # If the save operation failed due to read-only, prompt user for a new file path
            while not success:
                new_output_path = filedialog.asksaveasfilename(title="Save Output Origin File As",
                                                               filetypes=[("Origin Project Files", "*.opju")])
                if not new_output_path:
                    print("Save operation canceled by user.")
                    op.exit()
                    break  # Exit the loop if the user cancels the save operation

                new_output_path = os.path.abspath(new_output_path)
                success = save_origin_project(new_output_path)

                if success:
                    message = f"Processing completed! Output saved as:\n{new_output_path}"
                    messagebox.showinfo("Processing Complete", message)
                    op.exit()
                    break  # Exit the loop if the new save is successful
                else:
                    # Display an error message if the new save location is also read-only
                    messagebox.showerror("Error",
                                         "Selected save location is read-only. Please choose a different location.")
                    continue  # Continue the loop and prompt the user again

        except Exception as e:
            print(f"An error occurred: {str(e)}")
            op.exit()

    add_graphs_to_project_action()


root = tk.Tk()
root.withdraw()


# End of Code snippet 4: Add graphs to existing Origin project

# Code snippet 5: Function to exit the application
def exit_application():
    try:
        window.quit()  # Close the main GUI window, which ends the tkinter event loop
    except Exception as e:
        print(f"An error occurred while closing Origin: {str(e)}")


# End of Code snippet 5: Function to exit the application

# Create the main GUI window
window = tk.Tk()
window.title("FTIR Data Processing")
window.geometry("400x400")

# Create a frame for the header
header_frame = tk.Frame(window, padx=20, pady=20)
header_frame.pack()

# Create a label for the header
header_label = tk.Label(header_frame, text="FTIR Data Processing", font=("Helvetica", 16, "bold"))
header_label.pack()

# Create a frame for the content
content_frame = tk.Frame(window, padx=20, pady=20)
content_frame.pack()

# Step 1 Section
label_step1 = tk.Label(content_frame, text="Step 1: Rename the header with a CV voltage range.")
label_step1.pack(anchor="w")

rename_columns_button = tk.Button(content_frame, text="Rename Columns", command=rename_columns, bg="light blue")
rename_columns_button.pack(pady=5, anchor="w")

# Step 2 Section
label_step2 = tk.Label(content_frame, text="Step 2: Change the background spectrum with a chosen column.")
label_step2.pack(anchor="w")

process_background_data_button = tk.Button(content_frame, text="Reprocess Background", command=bg_processing,
                                           bg="light blue")
process_background_data_button.pack(pady=5, anchor="w")

# Create labels for step instructions
step3_label = tk.Label(content_frame, text="Step 3: Create Origin Project to add Graphs")
step3_label.pack(anchor="w")  # Align label to the left

# Create a button for Step 3 functionality
create_origin_graphs_button = tk.Button(content_frame, text="Create Origin Project",
                                        command=create_origin_graphs, bg="light blue")
create_origin_graphs_button.pack(pady=5, anchor="w")

# Create a label for Step 4
step4_label = tk.Label(content_frame, text="Step 4: Add Graphs to Existing Origin Project")
step4_label.pack(anchor="w")

# Create a button for Step 4 functionality
add_to_project_button = tk.Button(content_frame, text="Existing Origin Project",
                                  command=add_graphs_to_project, bg="light blue")
add_to_project_button.pack(pady=5, anchor="w")

# Exit Section
exit_label = tk.Label(content_frame, text="To quit, click below")
exit_label.pack(anchor="w")

exit_button = tk.Button(content_frame, text="Exit Application", command=exit_application, bg="red")
exit_button.pack(pady=10, anchor="w")

# Start the tkinter event loop
window.mainloop()
