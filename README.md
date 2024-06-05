# Excel Fuzzy Matching Script

This project provides a Node.js script to perform fuzzy matching between names from two different Excel files. The script uses the `fuzzyset.js` library to find approximate matches between names and update their status accordingly.

## Features

- Clean and normalize names for comparison.
- Perform fuzzy matching to identify approximate matches between names in two Excel files.
- Update the status of names based on matching results.
- Export the results to new Excel files.

## Prerequisites

- Node.js (v12 or later)
- npm (Node Package Manager)

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/yourusername/excel-fuzzy-matching.git
    cd excel-fuzzy-matching
    ```

2. **Install dependencies:**

    ```bash
    npm install
    ```

3. **Create a `.env` file in the root directory with the following content:**

    ```plaintext
    SORTANT=sortant
    ENTRANT=entrant
    ADMIS=admis

    FILES_1=C:\\Users\\M\\Downloads\\SS\\Fichier1.xlsx
    FILES_2=C:\\Users\\M\\Downloads\\SS\\Fichier2.xlsx
    ```

## Usage

1. **Prepare your Excel files and place them at the paths specified in the `.env` file.**

2. **Run the script:**

    ```bash
    node script.js
    ```

3. **Check the output files `fichierSortant.xlsx` and `fichierEntrant.xlsx` in the project directory.**

## Script Overview

### `script.js`

The main script file that performs the following tasks:

- Load environment variables using `dotenv`.
- Define functions to clean names and compare them using fuzzy matching.
- Load and parse Excel files using `xlsx`.
- Perform fuzzy matching between names from two Excel sheets.
- Update the status of names based on matching results.
- Export the updated data to new Excel files.

## Contributing

Contributions are welcome! Please follow the steps below to contribute:

1. **Fork the repository.**

2. **Create a new branch:**

    ```bash
    git checkout -b feature/your-feature-name
    ```

3. **Make your changes and commit them:**

    ```bash
    git commit -m 'Add some feature'
    ```

4. **Push to the branch:**

    ```bash
    git push origin feature/your-feature-name
    ```

5. **Open a pull request.**

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [xlsx](https://www.npmjs.com/package/xlsx) for parsing and writing Excel files.
- [fuzzyset.js](https://www.npmjs.com/package/fuzzyset.js) for fuzzy string matching.
- [dotenv](https://www.npmjs.com/package/dotenv) for managing environment variables.

# Contributing to Excel Fuzzy Matching Script

Thank you for considering contributing to our project! Here are some guidelines to help you get started.

## How to Contribute

### Reporting Bugs

If you find a bug, please report it by opening an issue on our GitHub repository. Provide as much detail as possible, including steps to reproduce the bug and any relevant error messages.

### Suggesting Enhancements

We welcome suggestions for new features and improvements. Please open an issue on our GitHub repository and describe your suggestion in detail.

### Submitting Pull Requests

1. **Fork the repository:**

    Click the "Fork" button at the top right corner of the repository page.

2. **Clone your forked repository:**

    ```bash
    git clone https://github.com/msbouady/fuzzyName-Macthing-System.git
    cd excel-fuzzy-matching
    ```

3. **Create a new branch:**

    ```bash
    git checkout -b feature/your-feature-name
    ```

4. **Make your changes:**

    Implement your changes, ensuring that your code follows the project's coding standards.

5. **Commit your changes:**

    ```bash
    git commit -m 'Add some feature'
    ```

6. **Push to the branch:**

    ```bash
    git push origin feature/your-feature-name
    ```

7. **Open a pull request:**

    Go to the original repository and click on the "New pull request" button. Provide a clear and descriptive title and description for your pull request.

### Coding Standards

- Follow the existing coding style.
- Write clear and descriptive commit messages.
- Ensure that your code is properly commented.

### Running Tests

Before submitting your pull request, make sure your changes do not break any existing functionality. Run the script and check the output files to ensure everything works as expected.

## Code of Conduct

We expect all contributors to adhere to the [Contributor Covenant Code of Conduct](https://www.contributor-covenant.org/version/2/0/code_of_conduct/). Please read it to understand the expectations for interactions in our community.



