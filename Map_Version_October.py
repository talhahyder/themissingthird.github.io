import pandas as pd
import folium
from folium.plugins import MarkerCluster

# Load the data from the Excel file
file_path = 'data.xlsx'
data = pd.read_excel(file_path)

# Extract relevant columns
tehsil = data.iloc[:, 0]  # Tehsil names (1st column)
children_out_of_school = data.iloc[:, 6]  # Number of children out of school (7th column)
province = data.iloc[:, 5]  # Province (6th column)
district = data.iloc[:, 4]  # District (5th column)
male = data.iloc[:, 7]  # Male (8th column)
female = data.iloc[:, 8]  # Female (9th column)
rural = data.iloc[:, 9]  # Rural (10th column)
urban = data.iloc[:, 12]  # Urban (13th column)

# Load Latitude and Longitude from the Excel file
data[['Latitude', 'Longitude']] = data.iloc[:, [1, 2]]  # Columns 2 and 3

# Filter out rows where coordinates are still missing
data = data.dropna(subset=['Latitude', 'Longitude'])

# Create a map centered on the first location with available coordinates
if not data.empty:
    # Use a custom tile for better aesthetics
    m = folium.Map(location=[data['Latitude'].mean(), data['Longitude'].mean()], zoom_start=6, tiles='cartodb positron')

    # Create a fixed heading with CSS styling
    heading_html = """
    <div style="
        position: fixed;
        top: 10px;
        width: 100%;
        font-size: 24px;
        font-weight: bold;
        text-align: center;
        background-color: rgba(255, 255, 255, 0.8);
        padding: 10px;
        z-index: 9999;
        font-family: Arial;">
    Visualization of Out of School Children in Pakistan
    </div>
    """
    m.get_root().html.add_child(folium.Element(heading_html))

    # Create a dropdown list for selecting Province
    province_dropdown_html = '''
    <div style="position: fixed; top: 50px; left: 10px; z-index: 9999; background-color: white; padding: 10px;">
        <label for="province-dropdown">Select Province:</label>
        <select id="province-dropdown" onchange="filterTehsilByProvince()">
            <option value="">Select Province</option>
    '''

    # Get unique province names and populate the dropdown
    unique_provinces = data['Province'].unique()
    for prov in unique_provinces:
        province_dropdown_html += f'<option value="{prov}">{prov}</option>'

    province_dropdown_html += '</select></div>'

    # Add the province dropdown to the map
    m.get_root().html.add_child(folium.Element(province_dropdown_html))

    # Create a dropdown list for selecting Tehsil
    tehsil_dropdown_html = '''
    <div style="position: fixed; top: 100px; left: 10px; z-index: 9999; background-color: white; padding: 10px;">
        <label for="tehsil-dropdown">Select Tehsil:</label>
        <select id="tehsil-dropdown" onchange="zoomToTehsil()">
            <option value="">Select Tehsil</option>
    '''

    # Populate the tehsil dropdown with all Tehsil names and coordinates initially
    for tehsil_name, lat, lon, prov in zip(tehsil, data['Latitude'], data['Longitude'], province):
        tehsil_dropdown_html += f'<option value="{lat},{lon}" data-province="{prov}">{tehsil_name}</option>'

    tehsil_dropdown_html += '</select></div>'

    # Add the tehsil dropdown to the map
    m.get_root().html.add_child(folium.Element(tehsil_dropdown_html))

    # JavaScript function to zoom to the selected Tehsil
    zoom_script = '''
    <script>
    function zoomToTehsil() {
        var tehsilDropdown = document.getElementById("tehsil-dropdown");
        var coords = tehsilDropdown.value.split(',');
        if (coords.length === 2) {
            var lat = parseFloat(coords[0]);
            var lon = parseFloat(coords[1]);
            map.setView([lat, lon], 10);  // Zoom level 10 for Tehsil
        }
    }

    function filterTehsilByProvince() {
        var provinceDropdown = document.getElementById("province-dropdown").value;
        var tehsilDropdown = document.getElementById("tehsil-dropdown");
        for (var i = 0; i < tehsilDropdown.options.length; i++) {
            var option = tehsilDropdown.options[i];
            if (option.getAttribute('data-province') === provinceDropdown || provinceDropdown === "") {
                option.style.display = "";  // Show the option
            } else {
                option.style.display = "none";  // Hide the option
            }
        }
        tehsilDropdown.selectedIndex = 0;  // Reset the tehsil dropdown to 'Select Tehsil'
    }
    </script>
    '''
    m.get_root().html.add_child(folium.Element(zoom_script))

    # Create a MarkerCluster object for grouping the markers
    marker_cluster = MarkerCluster().add_to(m)

    # Iterate through the data and create markers with enhanced popups
    for lat, lon, children, tehsil_name, prov, dist, male_count, female_count, rural_count, urban_count in zip(
        data['Latitude'], data['Longitude'], children_out_of_school, tehsil,
        province, district, male, female, rural, urban
    ):
        if pd.notna(lat) and pd.notna(lon):  # Check for NaN values
            # Info popup when clicking on the marker
            popup_info = f"""
            <b>Tehsil:</b> {tehsil_name} <br>
            <b>Province:</b> {prov} <br>
            <b>District:</b> {dist} <br>
            <b>Male:</b> {male_count:,} <br>
            <b>Female:</b> {female_count:,} <br>
            <b>Rural:</b> {rural_count:,} <br>
            <b>Urban:</b> {urban_count:,} <br>
            <b>Number of Children out of School:</b> {children:,}<br>
            <a href="javascript:void(0)" onclick="showDetails('{tehsil_name}', {male_count}, {female_count}, {rural_count}, {urban_count}, {children})">View Details</a>
            """

            folium.Marker(
                location=[lat, lon],
                popup=folium.Popup(popup_info, max_width=300),
            ).add_to(marker_cluster)

    # Save the map to an HTML file
    m.save('out_of_school_map.html')
    print("Map created and saved as 'out_of_school_map.html'")

    # Additional HTML and JavaScript for the details section
    with open('out_of_school_map.html', 'a') as f:
        f.write('''
        <div id="details" style="position: absolute; top: 100px; left: 10px; width: 350px; 
                                  background-color: white; border: 1px solid black; padding: 10px;
                                  display: none; z-index: 9999;">
            <h4 id="details-title"></h4>
            <div id="details-content"></div>
            <div id="charts"></div>
            <div id="female-distribution" style="margin-top: 10px;"></div>
        </div>
        
        <!-- Always Visible Disclaimer Section -->
        <div id="disclaimer" style="position: fixed; bottom: 10px; right: 10px; 
                                     background-color: rgba(255, 255, 255, 0.9); 
                                     padding: 5px; border-radius: 5px; 
                                     box-shadow: 0 0 5px rgba(0, 0, 0, 0.3);
                                     z-index: 10000;"> <!-- Adjusted z-index -->
            <a href="https://mathsandscience.pk/publications/the-missing-third/" 
               style="text-decoration: none; color: black; font-family: Arial;">
                Pak Alliance for Maths and Science
            </a>
            <br>
            <img src="https://mathsandscience.pk/wp-content/uploads/2021/01/cropped-PAMSM-Logo.jpg" 
                 alt="Pak Alliance for Maths and Science Logo" 
                 style="width: 50px; height: auto;">
        </div>

        <script>
        function showDetails(tehsil, male, female, rural, urban, children) {
            document.getElementById('details').style.display = 'block';
            document.getElementById('details-title').innerHTML = tehsil;
            document.getElementById('details-content').innerHTML = 
                '<b>Male:</b> ' + male.toLocaleString() + '<br>' +
                '<b>Female:</b> ' + female.toLocaleString() + '<br>' +
                '<b>Rural:</b> ' + rural.toLocaleString() + '<br>' +
                '<b>Urban:</b> ' + urban.toLocaleString() + '<br>' +
                '<b>Number of Children out of School:</b> ' + children.toLocaleString();
            createCharts(male, female, rural, urban, children);
            showFemaleDistribution(female, rural, urban);
        }

        function createCharts(male, female, rural, urban, children) {
            const chartData = {
                labels: ['Male', 'Female'],
                datasets: [{
                    label: 'Children Statistics',
                    data: [male, female],
                    backgroundColor: ['#42A5F5', '#FF7F50'] // Two-color palette
                }]
            };
            const pieData = {
                labels: ['Rural', 'Urban'],
                datasets: [{
                    label: 'Population Distribution',
                    data: [rural, urban],
                    backgroundColor: ['#42A5F5', '#FF7F50'] // Two-color palette
                }]
            };

            const ctxBar = document.createElement('canvas');
            const ctxPie = document.createElement('canvas');
            document.getElementById('charts').innerHTML = ''; // Clear previous charts
            document.getElementById('charts').appendChild(ctxBar);
            document.getElementById('charts').appendChild(ctxPie);

            new Chart(ctxBar.getContext('2d'), {
                type: 'bar',
                data: chartData,
                options: {
                    responsive: true,
                    scales: {
                        ...
                    }
                }
            });
        }
        </script>
        ''')
