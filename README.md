# Vaccination Data API

This API offers data about Covid-19-vaccinations in Germany and its federal states.
The data was fetched from the Robert Koch Institute (RKI). It offers a table with reported vaccinations nationwide and by federal state as well as by STIKO indication as an excel-file (https://www.rki.de/DE/Content/InfAZ/N/Neuartiges_Coronavirus/Daten/Impfquotenmonitoring.xlsx).
For the API the excel data was converted to a JSON object providing data for all federal states and in total.

## Technologies

The REST API was built with Node.js and Express.

## How to Use

The API is hosted under: https://www.rki.de/DE/Content/InfAZ/N/Neuartiges_Coronavirus/Daten/Impfquotenmonitoring.xlsx

The defined endpoint for the data is /vaccinations.
