<html>

<head>
    <title>Excel Tools</title>
    <link rel="stylesheet" href="vendor/bootstrap/css/bootstrap.css">
</head>

<body>
    <div class="container">
        <div class="jumbotron">
            <h1 class="display-4">Excel Tools</h1>
            <p class="lead">Export Import Tabel Warga</p>
            <hr class="my-4">
        </div>
        <div class="row">

            <div class="card col-md-6">
                <div class="card-body">
                    <h5 class="card-title">Form Import Excel</h5>
                    <form method="post" enctype="multipart/form-data" action="proses/import.php">
                        <div class="form-group">
                            <label for="exampleInputFile">File Upload</label>
                            <input type="file" name="berkas_excel" class="form-control-file" id="exampleInputFile">
                        </div>
                        <button type="submit" class="btn btn-primary">Import</button>
                    </form>
                </div>
            </div>

            <div class="card col">
                <div class="card-body">
                    <h5 class="card-title">Form Export Excel</h5>

                    <p>File hasil export = htdocs/Data Export/<strong>Data Warga.xlsx</strong></p>

                    <div class="alert alert-warning p-0" role="alert">
                        Data Export otomatis replace !!!
                    </div>

                    <p></p>
                    <a class="btn btn-success" href="proses/export.php">Export</a>
                </div>
            </div>

        </div>
    </div>
</body>

</html>

<script src="vendor/bootstrap/js/bootstrap.js"></script>