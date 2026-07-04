# DataMiner Pro

DataMiner Pro is a C# Windows Forms application designed to collect business lead data from Google Places using an external API.
The application allows users to search for businesses by city, country, and business category, then displays the results in a table and exports them to an Excel file.

## Project Idea

The main idea of this project is to help users collect useful business information such as business names, phone numbers, ratings, cities, and websites.
This can be useful for lead generation, market research, and building contact lists for different business categories.

## Features

* Search businesses by city
* Select business category such as:

  * Restaurants
  * Hotels
  * Cafes
  * Clinics
  * Supermarkets
  * Pharmacies
* Get city coordinates automatically using Nominatim OpenStreetMap API
* Fetch business data using Apify Google Places API
* Display results in a DataGridView
* Filter businesses by rating
* Remove duplicate business names
* Export results to Excel file
* Save phone numbers as text in Excel to avoid formatting problems

## Technologies Used

* C#
* Windows Forms
* .NET Framework
* HttpClient
* Newtonsoft.Json
* ClosedXML
* Nominatim OpenStreetMap API
* Apify API

## Data Collected

The application collects and displays the following information:

* Business Name
* Phone Number
* City
* Rating
* Website

## How It Works

1. The user enters the city name.
2. The application gets the latitude and longitude of the city using Nominatim API.
3. The user selects the business category.
4. The application sends a request to the Apify Google Places API.
5. The returned data is filtered and displayed in the table.
6. The user can export the final results to an Excel file.

## Requirements

Before running the project, make sure you have:

* Visual Studio installed
* .NET Framework support
* Internet connection
* Apify API token
* Required NuGet packages:

  * Newtonsoft.Json
  * ClosedXML

## Important Note

For security reasons, do not upload your real API key to GitHub.

Instead of writing the API key directly in the code, it is better to store it in a safe place such as:

* App.config
* Environment variable
* Local settings file ignored by Git

Example:

```csharp
string apiKey = "YOUR_API_KEY_HERE";
```

## Project Status

This project is a learning and practical project built to improve skills in:

* C# Windows Forms
* Working with APIs
* JSON parsing
* Data filtering
* Excel export
* Building real-world desktop applications

## Screenshots

 

```md
<img width="1517" height="682" alt="image" src="https://github.com/user-attachments/assets/c16d2523-1a7d-4028-9e1c-b1a6290871ef" />

 <img width="1092" height="625" alt="image" src="https://github.com/user-attachments/assets/b3fc32f9-8f9f-4265-b7c5-450378647f8e" />

```

## Future Improvements

Possible future improvements:

* Add more search filters
* Add country validation
* Improve UI design
* Add loading progress bar
* Save previous searches
* Add database storage
* Add email and social media extraction if available from the API
* Support multiple cities in one search

## Author

Developed by Yaseen Hani Abdul-Majeed Eimirat

Computer Systems Engineering Student
