#DYNAMIC LUSMATRIX XLS FILE TO YEARLY LUSMATRIX TXT FILES

#Created by Bep Schrammeijer
#July 2016
#Creates lusmatrix files for each year

#packages required
require(openxlsx)

#' Function to generate a lusmatrix file from existing Excel data
#'
#' [more detailed description what the function does here]
#'
#'
#' @param workbook Workbook object containing the data.
#' @param sheet_name Character name of the sheet where the data is
#' @param rows Numeric vector defining the row range of cells processed.
#' @param cols Numeric vector defining the column range of cells processed.
#' @param ouput_dir Character path to output directory.
#'
#' @return invisible logical (TRUE if successful). The function is used
#'         only for it's side effects (processing and writing data).

process_year <- function(woorkbook, sheet_name, rows, cols,
                         output_dir) {
  # Read in the data. Pass on the function deafault arguments, leave
  # others hard coded (e.g. col and row names)
  ls_matrix <- readWorkbook(woorkbook, sheet = sheet_name,
                            colNames = FALSE, rowNames = FALSE,
                            rows = rows, cols = cols)
  # Round the columns
  ls_matrix[,1:3] <- round(ls_matrix[,1:3], 0)
  ls_matrix[,4] <- round(ls_matrix[,4], 2)
  # Dynamically generate the file name. We could provide the number
  # of the year as an argument, but we can also read it from the
  # Excel file since it's always in the same place. Use the first
  # elements of the rows and cols vectors.
  year_no_row <- rows[1] - 3
  year_no_col <- cols[1]
  year_no <- readWorkbook(woorkbook, sheet = sheet_name,
                          colNames = FALSE, rowNames = FALSE,
                          rows = year_no_row, cols = year_no_col)
  # Object year_no is still a data.frame, let's extract the first
  # (and only) atomic element to get a single number. Note
  # that we are overwriting the previous object.
  year_no <- year_no[1, 1]
  # Let's give some info for the user as well
  message("Processing year ", year_no, "...")
  # Construct the file name. Use file.path to create the path
  # in a platform independent way.
  output_file <- file.path(output_dir, paste0("lusmatrix.", year_no))
  write.table(ls_matrix, output_file, sep = "\t",
              row.names = FALSE, col.names = FALSE)

  # No need to worry about the exact meaning of this, but it's good
  # to return something.
  return(invisible(TRUE))
}

# Since we have defined a new function above, let's use it. First,
# let's define the needed input arguments:

# Define data source
excel_ds <- "Scenario_SSP1/IME/ME/ME/lus_matrix_ME_ls24.xlsx"
ls_matrix <- loadWorkbook(excel_ds)

# sheet_name is always the same
sheet_name <- "dynamic output"

# row IDs are always the same
rows <- 16:39

# column IDs are not the same. The structure in the Excel sheet
# is static: first column ID is 2, last 452. A new data section
# (year) start every 9th column (51 years all together).
# The following creates a vector c(2, 11, 20, ..., )
col_starts <- seq(2, 452, 9)
# Create columns stop indexes, it's col_start + 4. R is "vectorized"
# for many expressions, meaning that the following will add 4 to each
# element of col_starts
col_stops <- col_starts + 3
# Bind the starts and stops together, getting a matrix
cols <- cbind(col_starts, col_stops)

# Define the output directory, always the same:
output_dir <- "Scenario_SSP1/IME/ME/ME/"

# We're good to go! Let's call our function with the defined arguments.
# Let's loop over the rows in cols:

for (i in 1:nrow(cols)) {
  # Get the current cols. Create a sequence from the start and stop
  # positions defined in cols
  current_cols <- seq(cols[i, 1], cols[i, 2], 1)
  process_year(woorkbook = ls_matrix, sheet_name = sheet_name,
               rows = rows, current_cols, output_dir)
}

message("All done!")
