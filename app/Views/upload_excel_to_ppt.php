<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Upload Excel to PowerPoint</title>
</head>
<body>
    <h2>Upload Excel File</h2>
    <form action="convert" method="POST" enctype="multipart/form-data">
        <input type="file" name="excel_file" accept=".xlsx">
        <button type="submit">Convert to PowerPoint</button>
    </form>
</body>
</html>
