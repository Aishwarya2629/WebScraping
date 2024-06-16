import requests
from bs4 import BeautifulSoup
import webbrowser
import tkinter as tk
from tkinter import ttk
from urllib.parse import urljoin
import pandas as pd

# Base URL of the site to scrape
base_url = "https://books.toscrape.com/"

# Function to get the soup object from a URL
def get_soup(url):
    response = requests.get(url)
    if response.status_code == 200:
        return BeautifulSoup(response.text, 'html.parser')
    else:
        print(f"Failed to retrieve the webpage. Status code: {response.status_code}")
        return None

# Get the main page soup
soup = get_soup(base_url)

# Find all genres
genres = soup.find('ul', class_='nav-list').find('ul').find_all('a')
genre_dict = {genre.text.strip(): urljoin(base_url, genre['href']) for genre in genres}

def on_genre_select(event):
    selected_genre = genre_combo.get()
    genre_url = genre_dict[selected_genre]
    genre_soup = get_soup(genre_url)

    # Find all books in the selected genre
    books = genre_soup.find_all('article', class_='product_pod')
    book_dict = {book.h3.a['title']: urljoin(genre_url, book.h3.a['href']) for book in books}

    # Clear the current book list
    for i in book_list.get_children():
        book_list.delete(i)

    # Populate the book list
    for book_title, book_url in book_dict.items():
        book_list.insert("", "end", text=book_title, values=(book_url,))

def on_book_select(event):
    selected_item = book_list.selection()[0]
    book_url = book_list.item(selected_item, "values")[0]
    book_details = get_book_details(book_url)
    save_to_excel(book_details)
    webbrowser.open(book_url)

def get_book_details(book_url):
    book_soup = get_soup(book_url)
    title = book_soup.find('div', class_='product_main').h1.text
    price = book_soup.find('p', class_='price_color').text
    availability = book_soup.find('p', class_='instock availability').text.strip()
    rating = book_soup.find('p', class_='star-rating')['class'][1]
    description = book_soup.find('meta', {'name': 'description'})['content'].strip()
    upc = book_soup.find('th', text='UPC').find_next_sibling('td').text

    details = {
        'Title': title,
        'Price': price,
        'Availability': availability,
        'Rating': rating,
        'Description': description,
        'UPC': upc,
        'URL': book_url
    }

    return details

def save_to_excel(book_details):
    df = pd.DataFrame([book_details])
    file_name = 'books.xlsx'
    
    try:
        # Try to read the existing file
        existing_df = pd.read_excel(file_name)
        # Append the new book details
        df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        # If the file does not exist, just use the new dataframe
        pass

    # Write the dataframe to the Excel file
    df.to_excel(file_name, index=False)
    print("Book details saved to books.xlsx")


# Create the main window
root = tk.Tk()
root.title("Book Scraper")

# Create and pack the genre selection combo box
genre_label = tk.Label(root, text="Select a genre:")
genre_label.pack(pady=5)
genre_combo = ttk.Combobox(root, values=list(genre_dict.keys()), state="readonly")
genre_combo.pack(pady=5)
genre_combo.bind("<<ComboboxSelected>>", on_genre_select)

# Create and pack the book list treeview
book_list = ttk.Treeview(root, columns=("URL"), show="tree")
book_list.pack(pady=10, padx=10, fill="both", expand=True)
book_list.bind("<Double-1>", on_book_select)

# Start the main loop
root.mainloop()
