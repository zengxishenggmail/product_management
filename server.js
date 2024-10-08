const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const { Sequelize, DataTypes } = require('sequelize');
const ExcelJS = require('exceljs');

const app = express();
const port = process.env.PORT || 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));
app.set('view engine', 'ejs');

let sequelize;

// Database connection
function connectToDatabase(config) {
  sequelize = new Sequelize(config.database, config.username, config.password, {
    host: config.host,
    port: config.port,
    dialect: 'mysql'
  });

  sequelize.authenticate()
    .then(() => {
      console.log('Connection has been established successfully.');
      initializeDatabase();
    })
    .catch(err => {
      console.error('Unable to connect to the database:', err);
    });
}

// Product model
let Product;

function initializeDatabase() {
  Product = sequelize.define('Product', {
    product_code: {
      type: DataTypes.STRING,
      allowNull: false,
      unique: true
    },
    product_name: {
      type: DataTypes.STRING,
      allowNull: false
    },
    qty: {
      type: DataTypes.INTEGER,
      allowNull: false
    },
    price: {
      type: DataTypes.DECIMAL(10, 2),
      allowNull: false
    },
    amount: {
      type: DataTypes.DECIMAL(10, 2),
      allowNull: false
    },
    remark: {
      type: DataTypes.TEXT
    },
    location: {
      type: DataTypes.STRING
    }
  });

  Product.sync({ force: true }).then(() => {
    console.log("Product table created");
  });
}

// Routes
app.get('/', (req, res) => {
  res.render('config');
});

app.post('/connect', (req, res) => {
  const config = {
    host: req.body.host,
    port: req.body.port,
    username: req.body.username,
    password: req.body.password,
    database: req.body.database
  };
  connectToDatabase(config);
  res.redirect('/products');
});

app.get('/products', async (req, res) => {
  const { product_name } = req.query;
  let where = {};
  if (product_name) {
    where.product_name = { [Sequelize.Op.like]: `%${product_name}%` };
  }
  const products = await Product.findAll({ where });
  res.render('products', { products, filter: product_name });
});

app.post('/products', async (req, res) => {
  const product = await Product.create(req.body);
  res.redirect('/products');
});

app.get('/products/:id', async (req, res) => {
  const product = await Product.findByPk(req.params.id);
  res.render('edit', { product });
});

app.post('/products/:id', async (req, res) => {
  await Product.update(req.body, { where: { id: req.params.id } });
  res.redirect('/products');
});

app.post('/products/:id/delete', async (req, res) => {
  await Product.destroy({ where: { id: req.params.id } });
  res.redirect('/products');
});

app.get('/export', async (req, res) => {
  const products = await Product.findAll();
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Products');
  
  worksheet.columns = [
    { header: 'Product Code', key: 'product_code', width: 15 },
    { header: 'Product Name', key: 'product_name', width: 25 },
    { header: 'Quantity', key: 'qty', width: 10 },
    { header: 'Price', key: 'price', width: 10 },
    { header: 'Amount', key: 'amount', width: 10 },
    { header: 'Remark', key: 'remark', width: 30 },
    { header: 'Location', key: 'location', width: 20 }
  ];

  products.forEach(product => {
    worksheet.addRow(product);
  });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=products.xlsx');

  await workbook.xlsx.write(res);
  res.end();
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});