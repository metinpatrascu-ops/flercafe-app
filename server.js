require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use(express.static('public'));

const HAS_OPENAI = !!process.env.OPENAI_API_KEY;
let openai = null;
if (HAS_OPENAI) {
  const OpenAI = require('openai');
  openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

const HAS_CLAUDE = !!process.env.ANTHROPIC_API_KEY;
let anthropic = null;
if (HAS_CLAUDE) {
  const Anthropic = require('@anthropic-ai/sdk');
  anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
}

// ─── MONGOOSE ──────────────────────────────────────────────────────────────
mongoose.connect(process.env.MONGODB_URI || 'mongodb://localhost:27017/flercafe')
  .then(() => console.log('MongoDB conectat'))
  .catch(err => console.error('MongoDB eroare conectare:', err.message));

const UserSchema = new mongoose.Schema({
  name: String,
  email: { type: String, unique: true },
  password: String,
  role: { type: String, default: 'owner' }
});
const User = mongoose.model('User', UserSchema);

const ProductSchema = new mongoose.Schema({
  name: { type: String, required: true },
  category: { type: String, enum: ['apa', 'soft', 'alcool', 'sirop', 'cafea', 'consumabile', 'altele'], default: 'altele' },
  unit: { type: String, default: 'buc' },
  packageSize: { type: Number, default: 1 }, // ex: 0.75 litri per sticla
  packageUnit: { type: String, default: 'L' },
  stockQuantity: { type: Number, default: 0 },
  minStock: { type: Number, default: 0 },
  purchasePrice: { type: Number, default: 0 }, // per unitate de pachet
  barcode: String,
  notes: String,
  active: { type: Boolean, default: true }
}, { timestamps: true });
const Product = mongoose.model('Product', ProductSchema);

const InvoiceSchema = new mongoose.Schema({
  date: { type: Date, default: Date.now },
  supplier: String,
  imagePath: String,
  rawOcrText: String,
  items: [{
    productId: { type: mongoose.Schema.Types.ObjectId, ref: 'Product' },
    productName: String,
    quantity: Number,
    unit: String,
    unitPrice: Number,
    totalPrice: Number
  }],
  totalAmount: Number,
  status: { type: String, enum: ['processing', 'done', 'manual'], default: 'manual' },
  notes: String
}, { timestamps: true });
const Invoice = mongoose.model('Invoice', InvoiceSchema);

const RecipeSchema = new mongoose.Schema({
  name: { type: String, required: true },
  category: { type: String, enum: ['cocktail', 'mocktail', 'cafea', 'bautura_calda', 'altele'], default: 'altele' },
  servingSizeMl: { type: Number, default: 250 },
  ingredients: [{
    productId: { type: mongoose.Schema.Types.ObjectId, ref: 'Product' },
    productName: String,
    quantity: Number,
    unit: String
  }],
  sellingPrice: { type: Number, default: 0 },
  notes: String,
  active: { type: Boolean, default: true }
}, { timestamps: true });
const Recipe = mongoose.model('Recipe', RecipeSchema);

const EventSchema = new mongoose.Schema({
  name: { type: String, required: true },
  client: String,
  date: Date,
  guestCount: { type: Number, default: 0 },
  durationHours: { type: Number, default: 4 },
  eventType: { type: String, enum: ['corporate', 'nunta', 'party', 'lansare', 'altele'], default: 'altele' },
  season: { type: String, enum: ['vara', 'iarna', 'primavara', 'toamna'], default: 'vara' },
  briefText: String,
  aiAnalysis: String,
  recommendations: [{
    productName: String,
    category: String,
    estimatedConsumptionL: Number,
    recommendedQuantity: Number,
    recommendedUnit: String,
    reason: String,
    warningMessage: String
  }],
  menuItems: [{
    name: String,
    quantity: Number,
    unit: String
  }],
  estimatedCost: { type: Number, default: 0 },
  offeredPrice: { type: Number, default: 0 },
  margin: { type: Number, default: 0 },
  status: { type: String, enum: ['draft', 'oferta', 'confirmat', 'finalizat'], default: 'draft' },
  notes: String,
  stockConsumed: [{
    productName: String,
    quantity: Number,
    unit: String,
    unitPrice: Number,
    totalCost: Number
  }],
  totalStockCost: { type: Number, default: 0 }
}, { timestamps: true });
const Event = mongoose.model('Event', EventSchema);

const StockMovementSchema = new mongoose.Schema({
  productId: { type: mongoose.Schema.Types.ObjectId, ref: 'Product', required: true },
  productName: String,
  type: { type: String, enum: ['in', 'out'], required: true },
  quantity: Number,
  reason: { type: String, enum: ['factura', 'eveniment', 'manual', 'pierdere'] },
  referenceId: String,
  referenceName: String,
  date: { type: Date, default: Date.now },
  notes: String
}, { timestamps: true });
const StockMovement = mongoose.model('StockMovement', StockMovementSchema);

// ─── MULTER ────────────────────────────────────────────────────────────────
const upload = multer({
  dest: 'public/uploads/',
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Doar imagini sunt acceptate'));
  }
});

const uploadExcel = multer({
  dest: 'public/uploads/',
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = file.mimetype.includes('spreadsheet') ||
                file.mimetype.includes('excel') ||
                file.originalname.match(/\.(xlsx|xls)$/i);
    if (ok) cb(null, true);
    else cb(new Error('Doar fișiere Excel (.xlsx, .xls) sunt acceptate'));
  }
});

// ─── AUTH MIDDLEWARE ────────────────────────────────────────────────────────
function verifyToken(req, res, next) {
  const auth = req.headers.authorization || (req.query.token ? `Bearer ${req.query.token}` : null);
  if (!auth) return res.status(401).json({ error: 'Token lipsă' });
  try {
    const decoded = jwt.verify(auth.replace('Bearer ', ''), process.env.JWT_SECRET);
    req.userId = decoded.userId;
    next();
  } catch {
    res.status(401).json({ error: 'Token invalid' });
  }
}

// ─── AUTH ROUTES ────────────────────────────────────────────────────────────
app.post('/api/auth/register', async (req, res) => {
  try {
    const { name, email, password } = req.body;
    if (!name || !email || !password) return res.status(400).json({ error: 'Toate câmpurile sunt obligatorii' });
    const exists = await User.findOne({ email });
    if (exists) return res.status(400).json({ error: 'Email-ul există deja' });
    const hashed = await bcrypt.hash(password, 10);
    const user = await User.create({ name, email, password: hashed });
    const token = jwt.sign({ userId: user._id }, process.env.JWT_SECRET, { expiresIn: '30d' });
    res.json({ token, user: { _id: user._id, name: user.name, email: user.email } });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    const user = await User.findOne({ email });
    if (!user || !await bcrypt.compare(password, user.password))
      return res.status(401).json({ error: 'Email sau parolă greșite' });
    const token = jwt.sign({ userId: user._id }, process.env.JWT_SECRET, { expiresIn: '30d' });
    res.json({ token, user: { _id: user._id, name: user.name, email: user.email } });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/auth/me', verifyToken, async (req, res) => {
  const user = await User.findById(req.userId).select('-password');
  res.json(user);
});

// ─── PRODUCTS ───────────────────────────────────────────────────────────────
app.get('/api/products', verifyToken, async (req, res) => {
  const products = await Product.find({ active: true }).sort({ category: 1, name: 1 });
  res.json(products);
});

app.post('/api/products', verifyToken, async (req, res) => {
  try {
    const product = await Product.create(req.body);
    res.json(product);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.put('/api/products/:id', verifyToken, async (req, res) => {
  try {
    const product = await Product.findByIdAndUpdate(req.params.id, req.body, { new: true });
    res.json(product);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.delete('/api/products/:id', verifyToken, async (req, res) => {
  await Product.findByIdAndUpdate(req.params.id, { active: false });
  res.json({ ok: true });
});

// ─── INVOICES ───────────────────────────────────────────────────────────────
app.post('/api/invoices/scan', verifyToken, upload.single('image'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Nicio imagine încărcată' });

    if (!HAS_OPENAI) {
      const invoice = await Invoice.create({
        imagePath: req.file.path.replace('public/', ''),
        date: new Date(),
        status: 'manual'
      });
      return res.json({
        invoice,
        extracted: { supplier: '', items: [], totalAmount: null },
        noAI: true,
        message: 'Nu este configurată o cheie OpenAI. Completează datele manual.'
      });
    }

    const imageBuffer = fs.readFileSync(req.file.path);
    const base64Image = imageBuffer.toString('base64');
    const mimeType = req.file.mimetype;

    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [{
        role: 'user',
        content: [
          {
            type: 'text',
            text: `Ești un expert contabil. Analizează această factură de la un bar/cafenea din România și extrage TOATE produsele.
Returnează STRICT un JSON valid cu această structură (fără text suplimentar):
{
  "supplier": "numele furnizorului",
  "date": "YYYY-MM-DD sau null",
  "totalAmount": 123.45,
  "items": [
    {
      "productName": "numele produsului",
      "quantity": 10,
      "unit": "bax/sticla/kg/L/buc",
      "unitPrice": 5.50,
      "totalPrice": 55.00
    }
  ]
}
Dacă nu poți citi clar o valoare, pune null. Produsele pot fi: ape, băuturi răcoritoare, alcool, cafea, siropuri, consumabile.`
          },
          {
            type: 'image_url',
            image_url: { url: `data:${mimeType};base64,${base64Image}`, detail: 'high' }
          }
        ]
      }],
      max_tokens: 2000
    });

    let extractedText = response.choices[0].message.content;
    extractedText = extractedText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

    let parsed;
    try {
      parsed = JSON.parse(extractedText);
    } catch {
      parsed = { error: 'Nu am putut parsa răspunsul AI', raw: extractedText };
    }

    const invoice = await Invoice.create({
      supplier: parsed.supplier,
      date: parsed.date ? new Date(parsed.date) : new Date(),
      imagePath: req.file.path.replace('public/', ''),
      rawOcrText: extractedText,
      items: (parsed.items || []).map(item => ({
        productName: item.productName,
        quantity: item.quantity,
        unit: item.unit,
        unitPrice: item.unitPrice,
        totalPrice: item.totalPrice
      })),
      totalAmount: parsed.totalAmount,
      status: 'done'
    });

    res.json({ invoice, extracted: parsed });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/invoices/confirm', verifyToken, async (req, res) => {
  try {
    const { invoiceId, items } = req.body;
    const invoice = await Invoice.findById(invoiceId);
    if (!invoice) return res.status(404).json({ error: 'Factură negăsită' });

    for (const item of items) {
      if (!item.productId) continue;

      await Product.findByIdAndUpdate(item.productId, {
        $inc: { stockQuantity: item.quantity },
        purchasePrice: item.unitPrice
      });

      await StockMovement.create({
        productId: item.productId,
        productName: item.productName,
        type: 'in',
        quantity: item.quantity,
        reason: 'factura',
        referenceId: invoice._id.toString(),
        referenceName: `Factură ${invoice.supplier || ''} ${new Date(invoice.date).toLocaleDateString('ro-RO')}`
      });
    }

    invoice.status = 'done';
    await invoice.save();

    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/invoices/manual', verifyToken, async (req, res) => {
  try {
    const { supplier, date, items, totalAmount, notes } = req.body;
    const invoice = await Invoice.create({
      supplier, date: date ? new Date(date) : new Date(),
      items, totalAmount, notes, status: 'manual'
    });

    for (const item of items) {
      if (!item.productId) continue;
      await Product.findByIdAndUpdate(item.productId, {
        $inc: { stockQuantity: item.quantity },
        purchasePrice: item.unitPrice
      });
      await StockMovement.create({
        productId: item.productId,
        productName: item.productName,
        type: 'in',
        quantity: item.quantity,
        reason: 'factura',
        referenceId: invoice._id.toString(),
        referenceName: `Factură manuală ${supplier || ''}`
      });
    }
    res.json(invoice);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get('/api/invoices', verifyToken, async (req, res) => {
  const invoices = await Invoice.find().sort({ date: -1 }).limit(50);
  res.json(invoices);
});

// ─── RECIPES ────────────────────────────────────────────────────────────────
app.get('/api/recipes', verifyToken, async (req, res) => {
  const recipes = await Recipe.find({ active: true }).sort({ category: 1, name: 1 });
  res.json(recipes);
});

app.post('/api/recipes', verifyToken, async (req, res) => {
  try {
    const recipe = await Recipe.create(req.body);
    res.json(recipe);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.put('/api/recipes/:id', verifyToken, async (req, res) => {
  const recipe = await Recipe.findByIdAndUpdate(req.params.id, req.body, { new: true });
  res.json(recipe);
});

app.delete('/api/recipes/:id', verifyToken, async (req, res) => {
  await Recipe.findByIdAndUpdate(req.params.id, { active: false });
  res.json({ ok: true });
});

// ─── EVENTS & AI ANALYSIS ───────────────────────────────────────────────────
app.get('/api/events', verifyToken, async (req, res) => {
  const events = await Event.find().sort({ date: -1 }).limit(50);
  res.json(events);
});

app.post('/api/events', verifyToken, async (req, res) => {
  try {
    const { menuItems, ...rest } = req.body;
    const event = await Event.create(rest);

    // Dacă există produse planificate → le salvăm direct ca stoc consumat cu prețuri
    if (menuItems && menuItems.length > 0) {
      const allProducts = await Product.find({ active: true }, 'name purchasePrice unit');
      const stockConsumed = menuItems.map(item => {
        const prod = allProducts.find(p =>
          p.name.toLowerCase() === item.name.toLowerCase() ||
          p.name.toLowerCase().includes(item.name.toLowerCase()) ||
          item.name.toLowerCase().includes(p.name.toLowerCase())
        );
        const unitPrice = prod?.purchasePrice || 0;
        const quantity = Number(item.quantity) || 0;
        return {
          productName: item.name,
          quantity,
          unit: item.unit || prod?.unit || 'buc',
          unitPrice,
          totalCost: Math.round(unitPrice * quantity * 100) / 100
        };
      });

      const totalStockCost = Math.round(stockConsumed.reduce((s, i) => s + i.totalCost, 0) * 100) / 100;
      const offeredPrice = Math.round(totalStockCost * 2.5 * 100) / 100;

      await Event.findByIdAndUpdate(event._id, { stockConsumed, totalStockCost, offeredPrice });

      return res.json({ ...event.toObject(), stockConsumed, totalStockCost, offeredPrice });
    }

    res.json(event);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.put('/api/events/:id', verifyToken, async (req, res) => {
  const event = await Event.findByIdAndUpdate(req.params.id, req.body, { new: true });
  res.json(event);
});

app.get('/api/events/:id', verifyToken, async (req, res) => {
  const event = await Event.findById(req.params.id);
  if (!event) return res.status(404).json({ error: 'Eveniment negăsit' });
  res.json(event);
});

// Salvează consumul efectiv al unui eveniment + creează mișcări de stoc
app.put('/api/events/:id/consumption', verifyToken, async (req, res) => {
  try {
    const { stockConsumed } = req.body;
    const event = await Event.findById(req.params.id);
    if (!event) return res.status(404).json({ error: 'Eveniment negăsit' });

    // Lookup purchase prices from DB and calculate costs automatically
    const allProducts = await Product.find({ active: true }, 'name purchasePrice');
    for (const item of stockConsumed) {
      if (!item.unitPrice || item.unitPrice === 0) {
        const prod = allProducts.find(p => p.name.toLowerCase().includes(item.productName.toLowerCase()) || item.productName.toLowerCase().includes(p.name.toLowerCase()));
        if (prod && prod.purchasePrice) {
          item.unitPrice = prod.purchasePrice;
          item.totalCost = Math.round(prod.purchasePrice * (item.quantity || 0) * 100) / 100;
        }
      } else {
        item.totalCost = Math.round((item.unitPrice || 0) * (item.quantity || 0) * 100) / 100;
      }
    }
    const totalStockCost = Math.round(stockConsumed.reduce((s, i) => s + (i.totalCost || 0), 0) * 100) / 100;

    // Șterge mișcările vechi legate de acest eveniment
    await StockMovement.deleteMany({ referenceId: req.params.id, reason: 'eveniment' });

    // Creează mișcări noi de stoc pentru fiecare produs consumat
    for (const item of stockConsumed) {
      if (!item.productName || !item.quantity) continue;
      const product = await Product.findOne({ name: { $regex: new RegExp(item.productName.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'i') }, active: true });
      if (product) {
        await StockMovement.create({
          productId: product._id,
          productName: product.name,
          type: 'out',
          quantity: item.quantity,
          reason: 'eveniment',
          referenceId: req.params.id,
          referenceName: event.name,
          date: event.date || new Date()
        });
      }
    }

    await Event.findByIdAndUpdate(req.params.id, {
      stockConsumed,
      totalStockCost,
      status: 'finalizat'
    });

    res.json({ ok: true, totalStockCost, stockConsumed });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Snapshot stoc la data evenimentului (calculat din mișcări)
app.get('/api/events/:id/stock-snapshot', verifyToken, async (req, res) => {
  try {
    const event = await Event.findById(req.params.id);
    if (!event) return res.status(404).json({ error: 'Eveniment negăsit' });

    const eventDate = event.date ? new Date(event.date) : new Date();
    const products = await Product.find({ active: true });

    // Calculează stocul la data evenimentului pe baza mișcărilor
    const snapshot = await Promise.all(products.map(async (p) => {
      const movsBefore = await StockMovement.find({
        productId: p._id,
        date: { $lte: eventDate }
      });
      const stockAtDate = movsBefore.reduce((sum, m) => {
        return sum + (m.type === 'in' ? m.quantity : -m.quantity);
      }, 0);

      return {
        productId: p._id,
        productName: p.name,
        category: p.category,
        unit: p.unit,
        purchasePrice: p.purchasePrice,
        stockAtEvent: stockAtDate > 0 ? stockAtDate : p.stockQuantity,
        currentStock: p.stockQuantity
      };
    }));

    res.json(snapshot.filter(s => s.stockAtEvent > 0));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Catalog de produse cunoscute cu reguli de detecție din textul clientului
const KNOWN_PRODUCTS = [
  // Apă — ordinea contează: 330ml înainte de generic
  { patterns: [/ap[aă]\s*plat[aă].*330|330.*ap[aă]\s*plat[aă]/i],
    name: 'Apă Plată Premium 330ml', category: 'apa',
    perPerson: 1.5, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  { patterns: [/ap[aă]\s*plat[aă]/i],
    name: 'Apă Plată Premium 750ml', category: 'apa',
    perPerson: 0.8, bottleL: 0.75, caseSize: 12, caseLabel: 'baxuri (12 sticle × 0.75L)' },
  { patterns: [/ap[aă]\s*mineral[aă].*330|330.*ap[aă]\s*mineral[aă]/i],
    name: 'Apă Minerală Premium 330ml', category: 'apa',
    perPerson: 1.0, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  { patterns: [/ap[aă]\s*mineral[aă]/i],
    name: 'Apă Minerală Premium 750ml', category: 'apa',
    perPerson: 0.5, bottleL: 0.75, caseSize: 12, caseLabel: 'baxuri (12 sticle × 0.75L)' },
  // Bere — 0% înainte de generic
  { patterns: [/heineken\s*0%|heineken\s*zero|bere\s*0%|bere.*f[aă]r[aă]\s*alcool/i],
    name: 'Heineken 0% Alcool 0.33L', category: 'soft',
    perPerson: 0.8, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  { patterns: [/heineken|bere\b/i],
    name: 'Bere Heineken 0.33L', category: 'alcool',
    perPerson: 1.8, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  // Soft
  { patterns: [/coca.?cola|cola\b|pepsi/i],
    name: 'Coca-Cola 0.33L', category: 'soft',
    perPerson: 0.8, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  { patterns: [/tonic[aă]|ap[aă]\s*tonic[aă]/i],
    name: 'Apă Tonică 0.33L', category: 'soft',
    perPerson: 0.5, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  { patterns: [/grapefruit\s*soda|soda.*grapefruit/i],
    name: 'Soda Grapefruit 0.33L', category: 'soft',
    perPerson: 0.4, bottleL: 0.33, caseSize: 24, caseLabel: 'baxuri (24 sticle × 0.33L)' },
  // Băuturi crafting
  { patterns: [/socat[aă]|sirop\s*soc\b/i],
    name: 'Socată / Sirop Soc', category: 'sirop',
    perPerson: 0.35, bottleL: 0.75, caseSize: 1, caseLabel: 'sticle 0.75L' },
  { patterns: [/sirop/i],
    name: 'Sirop Bar', category: 'sirop',
    perPerson: 0.05, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  // Alcool spirtos
  { patterns: [/prosecco|spumant|șampanie|sampanie/i],
    name: 'Prosecco Bortolin', category: 'alcool',
    perPerson: 1/6, bottleL: 0.75, caseSize: 1, caseLabel: 'sticle (1 sticlă = 6 pahare welcome)',
    welcomeDrink: true },
  { patterns: [/don\s*julio|tequila/i],
    name: 'Don Julio Tequila', category: 'alcool',
    perPerson: 0.04, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/johnnie|walker|whisky|whiskey/i],
    name: 'Johnnie Walker', category: 'alcool',
    perPerson: 0.04, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/bourbon|bulleit/i],
    name: 'Bulleit Bourbon', category: 'alcool',
    perPerson: 0.04, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/vodca|vodka|ketel/i],
    name: 'Ketel One Vodka', category: 'alcool',
    perPerson: 0.04, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/gin\b|tanqueray/i],
    name: 'Tanqueray Gin', category: 'alcool',
    perPerson: 0.04, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/aperol/i],
    name: 'Aperol', category: 'alcool',
    perPerson: 0.04, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/limoncello/i],
    name: 'Limoncello', category: 'alcool',
    perPerson: 0.03, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  { patterns: [/amaretto/i],
    name: 'Amaretto', category: 'alcool',
    perPerson: 0.03, bottleL: 0.7, caseSize: 1, caseLabel: 'sticle 0.7L' },
  // Cafea
  { patterns: [/cafea|espresso|coffee/i],
    name: 'Cafea', category: 'cafea',
    perPerson: 0.02, bottleL: 1, caseSize: 1, caseLabel: 'kg' },
  // Fructe
  { patterns: [/lime/i],
    name: 'Lime', category: 'consumabile',
    perPerson: 0.12, bottleL: 1, caseSize: 1, caseLabel: 'kg' },
  { patterns: [/l[aă]m[aâ]i/i],
    name: 'Lămâi', category: 'consumabile',
    perPerson: 0.1, bottleL: 1, caseSize: 1, caseLabel: 'kg' },
  { patterns: [/grapefruit/i],
    name: 'Grapefruit', category: 'consumabile',
    perPerson: 0.1, bottleL: 1, caseSize: 1, caseLabel: 'kg' },
  { patterns: [/portocal/i],
    name: 'Portocale', category: 'consumabile',
    perPerson: 0.1, bottleL: 1, caseSize: 1, caseLabel: 'kg' },
  { patterns: [/ment[aă]/i],
    name: 'Mentă', category: 'consumabile',
    perPerson: 0.01, bottleL: 1, caseSize: 1, caseLabel: 'kg' },
];

// Calculator inteligent — citește briefText și detectează produsele cerute
function ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems, briefText }) {
  const n = guestCount;
  const h = durationHours || 4;
  const isCorporate = ['corporate', 'lansare'].includes(eventType);
  const isSummer = season === 'vara';
  const MARGIN = 1.15;
  const durationMult = Math.max(1, h / 4);
  const waterMult = (isCorporate ? 1.2 : 1) * (isSummer ? 1.25 : 1);
  const iceMult = isSummer ? 1.3 : 1;

  // Textul de căutare — brief + menuItems combinate
  const searchText = [
    briefText || '',
    (menuItems || []).map(m => m.name).join(' ')
  ].join(' ').toLowerCase();

  const rec = [];
  const detectedNames = new Set();
  let proseccoQty = 0;
  let hasAlcohol = false;
  let hasWater = false;

  // Detectează produsele menționate explicit în brief
  for (const prod of KNOWN_PRODUCTS) {
    if (prod.patterns.some(p => p.test(searchText))) {
      if (detectedNames.has(prod.name)) continue;
      detectedNames.add(prod.name);

      const totalL = n * prod.perPerson * durationMult * MARGIN;
      const qty = prod.caseSize > 1
        ? Math.ceil(totalL / (prod.bottleL * prod.caseSize))
        : Math.ceil(totalL / prod.bottleL);

      if (prod.welcomeDrink) {
        // Prosecco: calcul pahare welcome
        const minSticle = Math.ceil(n / 6);
        proseccoQty = minSticle;
        const warn = qty < minSticle
          ? `⚠️ ${qty} sticle = ${qty * 6} pahare. Pentru ${n} invitați recomandăm minim ${minSticle} sticle.`
          : null;
        rec.push({
          productName: prod.name,
          category: prod.category,
          estimatedConsumptionL: null,
          recommendedQuantity: Math.max(qty, minSticle),
          recommendedUnit: prod.caseLabel,
          reason: `Welcome drink pentru ${n} invitați (1 sticlă = 6 pahare) + 15% marjă`,
          warningMessage: warn
        });
      } else {
        if (prod.category === 'apa') hasWater = true;
        if (prod.category === 'alcool') hasAlcohol = true;
        rec.push({
          productName: prod.name,
          category: prod.category,
          estimatedConsumptionL: +(totalL).toFixed(1),
          recommendedQuantity: qty,
          recommendedUnit: prod.caseLabel,
          reason: `${(prod.perPerson * durationMult).toFixed(2)}L/pers × ${n} inv. + 15% marjă`,
          warningMessage: null
        });
      }
    }
  }

  // Dacă nu s-a detectat nimic specific → adaugă pachete default
  if (rec.length === 0 || !hasWater) {
    const wL = n * 0.75 * waterMult * durationMult * MARGIN;
    rec.unshift({
      productName: 'Apă Plată Premium 750ml',
      category: 'apa',
      estimatedConsumptionL: +(wL).toFixed(1),
      recommendedQuantity: Math.ceil(wL / (0.75 * 12)),
      recommendedUnit: 'baxuri (12 sticle × 0.75L)',
      reason: `0.75L/pers × ${n} inv. ${isSummer ? '+ 25% vară' : ''} + 15% marjă`,
      warningMessage: null
    });
    const mL = n * 0.5 * waterMult * durationMult * MARGIN;
    rec.splice(1, 0, {
      productName: 'Apă Minerală Premium 750ml',
      category: 'apa',
      estimatedConsumptionL: +(mL).toFixed(1),
      recommendedQuantity: Math.ceil(mL / (0.75 * 12)),
      recommendedUnit: 'baxuri (12 sticle × 0.75L)',
      reason: `0.5L/pers × ${n} inv. ${isSummer ? '+ 25% vară' : ''} + 15% marjă`,
      warningMessage: null
    });
  }

  // Gheață — întotdeauna necesară
  const iceKg = Math.ceil((n * 0.4 + proseccoQty * 0.3) * iceMult * MARGIN);
  rec.push({
    productName: 'Gheață',
    category: 'consumabile',
    estimatedConsumptionL: null,
    recommendedQuantity: Math.ceil(iceKg / 5),
    recommendedUnit: `saci de 5kg (total ~${iceKg}kg)`,
    reason: `400g/pers bar${proseccoQty > 0 ? ` + 300g/sticlă prosecco` : ''}${isSummer ? ' + 30% vară' : ''} + 15% marjă`,
    warningMessage: isSummer ? 'Sezon cald — gheața se topește rapid. Asigură-te că ai congelator/ladă frigorifică.' : null
  });

  const warnings = [];
  if (isCorporate) warnings.push('Eveniment corporate/lansare: invitații vin cu mașina. Crește apa și soft-urile, reduci alcoolul.');
  if (isSummer) warnings.push('Temperaturi ridicate de vară: hidratarea este critică. Asigură-te că produsele sunt răcite înainte de eveniment.');

  const detectedList = [...detectedNames].join(', ');
  return {
    summary: `Eveniment ${eventType} cu ${n} invitați, ${h} ore, sezon ${season}.${detectedNames.size > 0 ? ` Produse detectate din brief: ${detectedList}.` : ' Pachete standard HoReCa.'} Cantitățile includ +15% marjă de siguranță.`,
    warnings,
    recommendations: rec,
    usedAI: false
  };
}

app.post('/api/events/analyze', verifyToken, async (req, res) => {
  try {
    const { guestCount, durationHours, eventType, season, briefText, menuItems } = req.body;
    if (!guestCount || guestCount < 1) return res.status(400).json({ error: 'Număr invitați invalid' });

    // Fără cheie API → calculator inteligent pe bază de reguli
    if (!HAS_CLAUDE && !HAS_OPENAI) {
      const analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems, briefText });
      return res.json({ analysis, usedAI: false });
    }

    // Claude este AI-ul principal; GPT-4o e fallback; reguli e fallback final
    const menuText = (menuItems || []).map(m => `- ${m.name}: ${m.quantity} ${m.unit}`).join('\n');

    const claudePrompt = `Ești un consultant expert în baruri de evenimente și HoReCa din România, cu 15 ani de experiență. Lucrezi pentru barul flērcafē.

BRIEF EVENIMENT:
- Client: ${clientName || 'nespecificat'}
- Invitați: ${guestCount} persoane
- Durată: ${durationHours || 4} ore
- Tip eveniment: ${eventType}
- Sezon: ${season}
- Ce dorește clientul (meniu/cerințe): "${briefText || 'nespecificat'}"
${menuText ? `- Produse specifice cerute:\n${menuText}` : ''}

SARCINA TA:
1. Citește CU ATENȚIE ce a scris clientul în câmpul "Ce dorește clientul"
2. Identifică EXACT produsele menționate și calculează cantitățile necesare pentru ${guestCount} invitați
3. Dacă clientul menționează "heineken" → calculează bax-uri de bere Heineken; dacă menționează "socată" → calculează sticle/litri de socată etc.
4. Adaugă și articolele critice care lipsesc (gheață, apă dacă nu e menționată)
5. Aplică standardele HoReCa: bere ~1.5-2 sticle/pers, apă ~0.75L/pers/4ore, gheață ~400g/pers

REGULI DE CALCUL:
- Apă plată/minerală 330ml: împărți la 0.33L per sticlă → nr. sticle, grupate în baxuri de 24
- Apă plată/minerală 750ml: împărți la 0.75L per sticlă → nr. sticle, baxuri de 12
- Bere 330ml: ~1.5-2 sticle/pers → baxuri de 24
- Bere 0% alcool: ~0.5-1 sticlă/pers
- Socată/crafting: ~0.35L/pers → nr. sticle de 750ml sau 1L
- Prosecco: 1 sticlă = 6 pahare de welcome
- Gheață: ~400g/pers + extra 30% vara
- Adaugă ÎNTOTDEAUNA +15% marjă de siguranță

RETURNEAZĂ STRICT JSON valid (fără text extra, fără markdown):
{"summary":"rezumat în 1-2 propoziții","warnings":["avertismente importante dacă există"],"recommendations":[{"productName":"numele exact al produsului","category":"apa|soft|alcool|sirop|cafea|consumabile|altele","estimatedConsumptionL":0.0,"recommendedQuantity":5,"recommendedUnit":"baxuri (24 sticle × 0.33L)","reason":"explicație calcul scurtă","warningMessage":null}],"usedAI":true}`;

    if (HAS_CLAUDE) {
      const msg = await anthropic.messages.create({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 2000,
        messages: [{ role: 'user', content: claudePrompt }]
      });
      let aiText = msg.content[0].text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      let analysis;
      try { analysis = JSON.parse(aiText); }
      catch { analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems, briefText }); }
      return res.json({ analysis, usedAI: true, aiProvider: 'claude' });
    }

    if (HAS_OPENAI) {
      const response = await openai.chat.completions.create({
        model: 'gpt-4o',
        messages: [{ role: 'user', content: claudePrompt }],
        max_tokens: 2000,
        temperature: 0.2
      });
      let aiText = response.choices[0].message.content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      let analysis;
      try { analysis = JSON.parse(aiText); }
      catch { analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems, briefText }); }
      return res.json({ analysis, usedAI: true, aiProvider: 'openai' });
    }

    res.json({ analysis: ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems, briefText }), usedAI: false });
  } catch (e) {
    try {
      const { guestCount, durationHours, eventType, season, menuItems, briefText } = req.body;
      const analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems, briefText });
      res.json({ analysis, usedAI: false });
    } catch (e2) {
      res.status(500).json({ error: e.message });
    }
  }
});

// ─── STOCK ──────────────────────────────────────────────────────────────────
app.post('/api/stock/import-excel', verifyToken, uploadExcel.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Niciun fișier încărcat' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(req.file.path);
    fs.unlinkSync(req.file.path);

    if (!wb.worksheets.length) return res.status(400).json({ error: 'Fișierul Excel nu conține niciun sheet' });

    const catMap = {
      'apă': 'apa', 'apa': 'apa', 'water': 'apa',
      'soft': 'soft', 'soft drinks': 'soft', 'cola': 'soft', 'suc': 'soft',
      'alcool': 'alcool', 'alcohol': 'alcool', 'vin': 'alcool', 'bere': 'alcool',
      'sirop': 'sirop', 'syrup': 'sirop',
      'cafea': 'cafea', 'coffee': 'cafea',
      'consumabile': 'consumabile', 'consumable': 'consumabile',
      'altele': 'altele', 'other': 'altele'
    };
    const normalizeCat = (val) => {
      if (!val) return 'altele';
      const v = val.toString().toLowerCase().trim();
      return catMap[v] || 'altele';
    };

    const eventTypeMap = {
      'corporate': 'corporate', 'lansare': 'lansare', 'launch': 'lansare',
      'nunta': 'nunta', 'nuntă': 'nunta', 'wedding': 'nunta',
      'party': 'party', 'festival': 'party', 'petrecere': 'party',
    };
    const normalizeEventType = (val) => {
      if (!val) return 'altele';
      const v = val.toString().toLowerCase().trim();
      for (const [k, mapped] of Object.entries(eventTypeMap)) {
        if (v.includes(k)) return mapped;
      }
      return 'altele';
    };

    const parseDate = (val) => {
      if (!val) return null;
      if (val instanceof Date) return val;
      const s = val.toString().trim();
      // dd.mm.yyyy / dd/mm/yyyy / yyyy-mm-dd
      const m1 = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
      if (m1) return new Date(`${m1[3]}-${m1[2].padStart(2,'0')}-${m1[1].padStart(2,'0')}`);
      const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m2) return new Date(s);
      const parsed = new Date(s);
      return isNaN(parsed) ? null : parsed;
    };

    const getCellVal = (row, col) => {
      const cell = row.getCell(col);
      if (!cell?.value) return '';
      if (cell.value?.text) return cell.value.text.toString().trim(); // rich text
      return cell.value.toString().trim();
    };
    const getCellNum = (row, col) => {
      const cell = row.getCell(col);
      if (!cell?.value) return 0;
      const v = parseFloat(cell.value.toString().replace(/[^\d.,-]/g, '').replace(',', '.'));
      return isNaN(v) ? 0 : v;
    };

    // ── detectare tip sheet ──────────────────────────────────────────────
    const isInventarSheet = (ws) => {
      let score = 0;
      ws.eachRow((row, rn) => {
        if (rn > 4) return;
        row.eachCell(cell => {
          const v = (cell.value || '').toString().toLowerCase();
          if (v.includes('produs') || v.includes('stoc') || v.includes('unitate')) score += 2;
          if (v.includes('preț') || v.includes('pret') || v.includes('categor')) score++;
        });
      });
      return score >= 2;
    };
    const isEvenimenteSheet = (ws) => {
      let score = 0;
      ws.eachRow((row, rn) => {
        if (rn > 4) return;
        row.eachCell(cell => {
          const v = (cell.value || '').toString().toLowerCase();
          if (v.includes('eveniment') || v.includes('event') || v.includes('client')) score += 2;
          if (v.includes('invita') || v.includes('persoane') || v.includes('data') || v.includes('dată')) score++;
        });
      });
      return score >= 2;
    };

    // ── import inventar ──────────────────────────────────────────────────
    const importInventar = async (ws) => {
      let headerRow = 1;
      let colMap = { name: 1, category: 2, stock: 3, unit: 4, price: 5, minStock: 7 };

      ws.eachRow((row, rowNum) => {
        if (rowNum > 5) return;
        row.eachCell((cell, colNum) => {
          const v = (cell.value || '').toString().toLowerCase().trim();
          if (v.includes('produs') || v === 'name' || v === 'denumire') { headerRow = rowNum; colMap.name = colNum; }
          if (v.includes('categor')) colMap.category = colNum;
          if (v.includes('stoc actual') || v === 'stoc' || v === 'stock quantity') colMap.stock = colNum;
          if (v === 'unitate' || v === 'unit' || v === 'um') colMap.unit = colNum;
          if ((v.includes('preț') || v.includes('pret') || v.includes('price')) && !v.includes('ofert')) colMap.price = colNum;
          if (v.includes('minim') || v.includes('min stock')) colMap.minStock = colNum;
        });
      });

      let created = 0, updated = 0, skipped = 0, errors = [];
      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const name = getCellVal(row, colMap.name);
        if (!name || name.toLowerCase().includes('total') || name.toLowerCase().includes('flērc') || name.toLowerCase().includes('flercafe')) continue;
        try {
          const data = {
            name, category: normalizeCat(getCellVal(row, colMap.category)),
            stockQuantity: getCellNum(row, colMap.stock),
            unit: getCellVal(row, colMap.unit) || 'buc',
            purchasePrice: getCellNum(row, colMap.price),
            minStock: getCellNum(row, colMap.minStock), active: true
          };
          const existing = await Product.findOne({ name: { $regex: new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}$`, 'i') }, active: true });
          if (existing) { await Product.findByIdAndUpdate(existing._id, data); updated++; }
          else { await Product.create(data); created++; }
        } catch (e) { errors.push(`Inv rând ${r}: ${e.message}`); skipped++; }
      }
      return { created, updated, skipped, errors };
    };

    // ── import evenimente ────────────────────────────────────────────────
    const importEvenimente = async (ws) => {
      let headerRow = 1;
      let colMap = { name: 1, client: 2, date: 3, guests: 4, type: 5, status: 6, price: 7, notes: 8 };

      ws.eachRow((row, rowNum) => {
        if (rowNum > 5) return;
        row.eachCell((cell, colNum) => {
          const v = (cell.value || '').toString().toLowerCase().trim();
          if (v.includes('eveniment') || v.includes('denumire') || v === 'event' || v === 'name') { headerRow = rowNum; colMap.name = colNum; }
          if (v === 'client' || v.includes('organizator') || v.includes('firma')) colMap.client = colNum;
          if (v.includes('data') || v.includes('dată') || v === 'date') colMap.date = colNum;
          if (v.includes('invitat') || v.includes('persoane') || v.includes('guests') || v.includes('nr.')) colMap.guests = colNum;
          if (v === 'tip' || v === 'type' || v.includes('tip event')) colMap.type = colNum;
          if (v === 'status' || v === 'stare') colMap.status = colNum;
          if (v.includes('preț') || v.includes('pret') || v.includes('ofert') || v.includes('price')) colMap.price = colNum;
          if (v.includes('note') || v.includes('observ') || v.includes('mentiu')) colMap.notes = colNum;
        });
      });

      const statusMap = {
        'draft': 'draft', 'oferta': 'oferta', 'ofertă': 'oferta', 'trimis': 'oferta',
        'confirmat': 'confirmat', 'confirmed': 'confirmat', 'da': 'confirmat',
        'finalizat': 'finalizat', 'done': 'finalizat', 'gata': 'finalizat'
      };
      const normalizeStatus = (v) => {
        if (!v) return 'draft';
        const s = v.toString().toLowerCase().trim();
        return statusMap[s] || 'draft';
      };

      let created = 0, updated = 0, skipped = 0, errors = [];
      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const name = getCellVal(row, colMap.name);
        if (!name || name.toLowerCase().includes('total')) continue;
        try {
          const eventDate = parseDate(ws.getRow(r).getCell(colMap.date)?.value);
          const data = {
            name,
            client: getCellVal(row, colMap.client),
            date: eventDate,
            guestCount: getCellNum(row, colMap.guests) || 0,
            eventType: normalizeEventType(getCellVal(row, colMap.type)),
            status: normalizeStatus(getCellVal(row, colMap.status)),
            offeredPrice: getCellNum(row, colMap.price),
            notes: getCellVal(row, colMap.notes),
            briefText: `Importat din Excel. Client: ${getCellVal(row, colMap.client)}`
          };
          const existing = await Event.findOne({ name: { $regex: new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}$`, 'i') } });
          if (existing) { await Event.findByIdAndUpdate(existing._id, data); updated++; }
          else { await Event.create(data); created++; }
        } catch (e) { errors.push(`Ev rând ${r}: ${e.message}`); skipped++; }
      }
      return { created, updated, skipped, errors };
    };

    // ── detectare sheet consum per eveniment ────────────────────────────
    const isConsumSheet = (ws) => {
      let score = 0;
      ws.eachRow((row, rn) => {
        if (rn > 4) return;
        row.eachCell(cell => {
          const v = (cell.value || '').toString().toLowerCase();
          if (v.includes('consum') || v.includes('folosit') || v.includes('utilizat')) score += 3;
          if (v.includes('eveniment') || v.includes('event')) score++;
          if (v.includes('produs') || v.includes('cantitat')) score++;
        });
      });
      return score >= 3;
    };

    // ── import consum per eveniment ──────────────────────────────────────
    const importConsum = async (ws) => {
      // Detectează coloanele: Eveniment | Produs | Cantitate | Unitate | Preț/buc
      let headerRow = 1;
      let colMap = { eventName: 1, product: 2, qty: 3, unit: 4, price: 5 };

      ws.eachRow((row, rowNum) => {
        if (rowNum > 5) return;
        row.eachCell((cell, colNum) => {
          const v = (cell.value || '').toString().toLowerCase().trim();
          if (v.includes('eveniment') || v === 'event') { headerRow = rowNum; colMap.eventName = colNum; }
          if (v.includes('produs') || v === 'product' || v === 'item') colMap.product = colNum;
          if (v.includes('cantit') || v === 'qty' || v === 'quantity') colMap.qty = colNum;
          if (v === 'unitate' || v === 'unit' || v === 'um') colMap.unit = colNum;
          if (v.includes('preț') || v.includes('pret') || v.includes('price')) colMap.price = colNum;
        });
      });

      // Grupează rândurile pe eveniment
      const byEvent = {};
      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        let evName = getCellVal(row, colMap.eventName);
        const product = getCellVal(row, colMap.product);
        if (!product) continue;
        if (!evName) evName = lastEvName || 'Necunoscut';
        else lastEvName = evName;

        if (!byEvent[evName]) byEvent[evName] = [];
        byEvent[evName].push({
          productName: product,
          quantity: getCellNum(row, colMap.qty),
          unit: getCellVal(row, colMap.unit) || 'buc',
          unitPrice: getCellNum(row, colMap.price),
          totalCost: getCellNum(row, colMap.qty) * getCellNum(row, colMap.price)
        });
      }

      let updated = 0;
      for (const [evName, consumed] of Object.entries(byEvent)) {
        const event = await Event.findOne({ name: { $regex: new RegExp(evName.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), 'i') } });
        if (!event) continue;
        const totalStockCost = consumed.reduce((s, i) => s + i.totalCost, 0);
        await Event.findByIdAndUpdate(event._id, { stockConsumed: consumed, totalStockCost, status: 'finalizat' });
        updated++;
      }
      return { updated, created: 0, skipped: 0, errors: [] };
    };

    let lastEvName = '';

    // ── procesează toate sheet-urile ─────────────────────────────────────
    const results = {
      produse: { created: 0, updated: 0, skipped: 0 },
      evenimente: { created: 0, updated: 0, skipped: 0 },
      consum: { updated: 0 },
      errors: []
    };

    for (const ws of wb.worksheets) {
      const sheetName = ws.name.toLowerCase();
      const looksLikeInv = isInventarSheet(ws);
      const looksLikeEv = isEvenimenteSheet(ws);
      const looksLikeConsum = isConsumSheet(ws);

      if (looksLikeConsum || sheetName.includes('consum') || sheetName.includes('folosit')) {
        const r = await importConsum(ws);
        results.consum.updated += r.updated;
        results.errors.push(...r.errors);
      } else if (sheetName.includes('eveniment') || sheetName.includes('event') || sheetName.includes('calendar') || (looksLikeEv && !looksLikeInv)) {
        const r = await importEvenimente(ws);
        results.evenimente.created += r.created;
        results.evenimente.updated += r.updated;
        results.evenimente.skipped += r.skipped;
        results.errors.push(...r.errors);
      } else if (looksLikeInv || sheetName.includes('inventar') || sheetName.includes('stoc') || sheetName.includes('stock')) {
        const r = await importInventar(ws);
        results.produse.created += r.created;
        results.produse.updated += r.updated;
        results.produse.skipped += r.skipped;
        results.errors.push(...r.errors);
      } else if (looksLikeEv) {
        const r = await importEvenimente(ws);
        results.evenimente.created += r.created;
        results.evenimente.updated += r.updated;
        results.evenimente.skipped += r.skipped;
        results.errors.push(...r.errors);
      }
    }

    const parts = [];
    if (results.produse.created || results.produse.updated)
      parts.push(`📦 Produse: ${results.produse.created} noi, ${results.produse.updated} actualizate`);
    if (results.evenimente.created || results.evenimente.updated)
      parts.push(`📅 Evenimente: ${results.evenimente.created} noi, ${results.evenimente.updated} actualizate`);
    if (results.consum.updated)
      parts.push(`🍹 Consum: ${results.consum.updated} evenimente actualizate cu date de consum`);
    if (!parts.length) parts.push('Nu am detectat date de importat. Verifică formatul fișierului.');

    res.json({ ok: true, message: parts.join(' · '), results });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/stock/export-excel', verifyToken, async (req, res) => {
  try {
    const products = await Product.find({ active: true }).sort({ category: 1, name: 1 });

    const wb = new ExcelJS.Workbook();
    wb.creator = 'flērcafē';
    wb.created = new Date();

    const ws = wb.addWorksheet('Inventar Stoc', { views: [{ state: 'frozen', ySplit: 2 }] });

    // Header principal
    ws.mergeCells('A1:G1');
    const titleCell = ws.getCell('A1');
    titleCell.value = `flērcafē — Inventar Stoc   |   ${new Date().toLocaleDateString('ro-RO')}`;
    titleCell.font = { bold: true, size: 14, color: { argb: 'FFC8A96E' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A1A' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 30;

    // Coloane
    ws.columns = [
      { header: 'Produs',          key: 'name',          width: 35 },
      { header: 'Categorie',       key: 'category',      width: 15 },
      { header: 'Stoc actual',     key: 'stock',         width: 14 },
      { header: 'Unitate',         key: 'unit',          width: 10 },
      { header: 'Preț / unitate',  key: 'price',         width: 16 },
      { header: 'Valoare stoc',    key: 'value',         width: 16 },
      { header: 'Stoc minim',      key: 'minStock',      width: 13 },
    ];

    // Stil header coloane (rândul 2)
    const headerRow = ws.getRow(2);
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FF111111' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8A96E' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FF888888' } } };
    });
    headerRow.height = 22;

    const catLabels = { apa: 'Apă', soft: 'Soft', alcool: 'Alcool', sirop: 'Sirop', cafea: 'Cafea', consumabile: 'Consumabile', altele: 'Altele' };
    let totalValue = 0;

    products.forEach((p, i) => {
      const val = p.stockQuantity * p.purchasePrice;
      totalValue += val;
      const isLow = p.minStock > 0 && p.stockQuantity <= p.minStock;
      const row = ws.addRow({
        name: p.name,
        category: catLabels[p.category] || p.category,
        stock: p.stockQuantity,
        unit: p.unit,
        price: p.purchasePrice,
        value: +val.toFixed(2),
        minStock: p.minStock || ''
      });
      row.height = 18;

      // Rânduri alternante
      const bg = i % 2 === 0 ? 'FFFAFAFA' : 'FFF3F3F3';
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'middle' };
      });

      // Format numeric
      row.getCell('price').numFmt = '#,##0.00 "RON"';
      row.getCell('value').numFmt = '#,##0.00 "RON"';
      row.getCell('stock').alignment = { horizontal: 'center', vertical: 'middle' };

      // Roșu dacă stoc scăzut
      if (isLow) {
        row.getCell('stock').font = { bold: true, color: { argb: 'FFCC0000' } };
        row.getCell('stock').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEEEE' } };
      }
    });

    // Rând total
    const totalRow = ws.addRow({ name: 'TOTAL VALOARE STOC', category: '', stock: '', unit: '', price: '', value: +totalValue.toFixed(2), minStock: '' });
    totalRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A1A' } };
      cell.font = { bold: true, color: { argb: 'FFC8A96E' } };
    });
    totalRow.getCell('value').numFmt = '#,##0.00 "RON"';
    totalRow.height = 22;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=flercafe-inventar-${new Date().toISOString().split('T')[0]}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/stock/summary', verifyToken, async (req, res) => {
  const products = await Product.find({ active: true });
  const lowStock = products.filter(p => p.stockQuantity <= p.minStock && p.minStock > 0);
  const totalValue = products.reduce((sum, p) => sum + (p.stockQuantity * p.purchasePrice), 0);
  res.json({ products, lowStock, totalValue, count: products.length });
});

app.post('/api/stock/movement', verifyToken, async (req, res) => {
  try {
    const { productId, type, quantity, reason, notes } = req.body;
    const product = await Product.findById(productId);
    if (!product) return res.status(404).json({ error: 'Produs negăsit' });

    const change = type === 'in' ? quantity : -quantity;
    await Product.findByIdAndUpdate(productId, { $inc: { stockQuantity: change } });

    const movement = await StockMovement.create({
      productId, productName: product.name, type, quantity, reason: reason || 'manual', notes
    });
    res.json(movement);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get('/api/stock/movements', verifyToken, async (req, res) => {
  const movements = await StockMovement.find().sort({ date: -1 }).limit(100);
  res.json(movements);
});

// ─── CALCULATOR ─────────────────────────────────────────────────────────────
app.post('/api/calculator/recipe-cost', verifyToken, async (req, res) => {
  try {
    const { ingredients } = req.body;
    let totalCost = 0;
    const breakdown = [];

    for (const ing of ingredients) {
      if (!ing.productId) continue;
      const product = await Product.findById(ing.productId);
      if (!product) continue;
      const cost = (product.purchasePrice / product.packageSize) * ing.quantity;
      totalCost += cost;
      breakdown.push({
        productName: product.name,
        quantity: ing.quantity,
        unit: ing.unit,
        unitCost: product.purchasePrice / product.packageSize,
        totalCost: cost
      });
    }

    res.json({ totalCost: Math.round(totalCost * 100) / 100, breakdown });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/calculator/event-offer', verifyToken, async (req, res) => {
  try {
    const { items, targetMarginPercent } = req.body;
    let totalCost = 0;
    const breakdown = [];

    for (const item of items) {
      const product = item.productId ? await Product.findById(item.productId) : null;
      const purchasePrice = product ? product.purchasePrice : (item.estimatedPrice || 0);
      const cost = purchasePrice * item.quantity;
      totalCost += cost;
      breakdown.push({
        name: item.name || product?.name || 'Necunoscut',
        quantity: item.quantity,
        unit: item.unit,
        purchasePrice,
        totalCost: cost
      });
    }

    const margin = targetMarginPercent || 200;
    const recommendedPrice = totalCost * (1 + margin / 100);

    res.json({
      totalCost: Math.round(totalCost * 100) / 100,
      recommendedPrice: Math.round(recommendedPrice * 100) / 100,
      marginPercent: margin,
      breakdown
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── DASHBOARD ──────────────────────────────────────────────────────────────
app.get('/api/dashboard', verifyToken, async (req, res) => {
  const [products, invoices, events, movements] = await Promise.all([
    Product.find({ active: true }),
    Invoice.find().sort({ createdAt: -1 }).limit(5),
    Event.find().sort({ date: -1 }).limit(5),
    StockMovement.find().sort({ date: -1 }).limit(10)
  ]);

  const lowStock = products.filter(p => p.minStock > 0 && p.stockQuantity <= p.minStock);
  const totalStockValue = products.reduce((s, p) => s + p.stockQuantity * p.purchasePrice, 0);
  const upcomingEvents = await Event.find({ date: { $gte: new Date() }, status: { $in: ['oferta', 'confirmat'] } }).sort({ date: 1 }).limit(3);

  res.json({ totalProducts: products.length, lowStock, totalStockValue, recentInvoices: invoices, recentEvents: events, upcomingEvents, recentMovements: movements });
});

// ─── AI CHAT ────────────────────────────────────────────────────────────────
function parseEventFromText(text) {
  const t = text;
  const tl = t.toLowerCase();

  const guestMatch = tl.match(/(\d+)\s*(?:persoan|invitat|oaspeț|guest)/i);
  const guestCount = guestMatch ? parseInt(guestMatch[1]) : 50;

  const hoursMatch = tl.match(/(\d+)\s*or[eă]/i);
  const durationHours = hoursMatch ? parseInt(hoursMatch[1]) : 4;

  const eventType = /corporate|business/i.test(tl) ? 'corporate' :
    /lansare/i.test(tl) ? 'lansare' :
    /nunt[aă]/i.test(tl) ? 'nunta' :
    /botez/i.test(tl) ? 'botez' : 'petrecere';

  const month = new Date().getMonth();
  const season = (month >= 4 && month <= 8) ? 'vara' : 'iarna';

  const menuItems = [];
  const prosMatch = tl.match(/(\d+)\s*sticle?\s*(?:de\s*)?prosecco/i);
  if (prosMatch) menuItems.push({ name: 'Prosecco', quantity: parseInt(prosMatch[1]), unit: 'sticle' });

  const nameMatch = t.match(/evenimentul?\s+([A-ZĂÂÎȘȚ][^\s,\.]+(?:\s+[A-ZĂÂÎȘȚ][^\s,\.]+)?)/);
  const eventName = nameMatch ? nameMatch[1] : 'Eveniment';

  const clientMatch = t.match(/client[ă]?\s+([A-ZĂÂÎȘȚ][^\s,\.]+)/i);
  const clientName = clientMatch ? clientMatch[1] : '';

  return { eventName, clientName, guestCount, durationHours, eventType, season, menuItems };
}

app.post('/api/ai/chat', verifyToken, async (req, res) => {
  try {
    const { message } = req.body;
    if (!message || !message.trim()) return res.status(400).json({ error: 'Mesaj gol' });

    let eventData, analysis, reply;

    const month = new Date().getMonth();
    const defaultSeason = (month >= 4 && month <= 8) ? 'vara' : 'iarna';

    const chatSystemPrompt = `Ești asistentul AI expert al barului flērcafē din România. Analizezi brief-uri de evenimente și calculezi necesarul de băuturi și produse. Comunici EXCLUSIV în română, prietenos și profesional.

Din mesajul utilizatorului extrage:
- Numele clientului/evenimentului
- Numărul de invitați
- Durata în ore (default 4)
- Tipul: corporate | lansare | nunta | botez | petrecere | altele
- Sezonul: vara | iarna (default: ${defaultSeason})
- Produsele menționate explicit cu cantitățile cerute

Returnează STRICT JSON valid:
{"eventName":"...","clientName":"...","guestCount":0,"durationHours":4,"eventType":"petrecere","season":"${defaultSeason}","menuItems":[{"name":"Heineken 0.33L","quantity":50,"unit":"sticle"}],"reply":"mesaj prietenos scurt că ai înțeles și urmează calculul"}`;

    if (HAS_CLAUDE) {
      const resp = await anthropic.messages.create({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 800,
        system: chatSystemPrompt,
        messages: [{ role: 'user', content: message }]
      });
      let txt = resp.content[0].text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      try { eventData = JSON.parse(txt); } catch { eventData = parseEventFromText(message); }
      reply = eventData.reply || `Am înțeles! Calculez pentru ${eventData.guestCount} invitați...`;
    } else if (HAS_OPENAI) {
      const resp = await openai.chat.completions.create({
        model: 'gpt-4o',
        messages: [{ role: 'system', content: chatSystemPrompt }, { role: 'user', content: message }],
        response_format: { type: 'json_object' },
        max_tokens: 700,
        temperature: 0.1
      });
      try { eventData = JSON.parse(resp.choices[0].message.content); } catch { eventData = parseEventFromText(message); }
      reply = eventData.reply || `Am înțeles! Calculez pentru ${eventData.guestCount} invitați...`;
    } else {
      eventData = parseEventFromText(message);
      reply = `Am înțeles! Calculez necesarul pentru ${eventData.guestCount} invitați, ${eventData.durationHours} ore.`;
    }

    analysis = ruleBasedAnalysis(eventData);
    res.json({ reply, eventData, analysis });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── PDF OFFER ───────────────────────────────────────────────────────────────
app.post('/api/offers/pdf', verifyToken, async (req, res) => {
  try {
    const PDFDocument = require('pdfkit');
    const { eventName, clientName, eventDate, guestCount, durationHours, items, notes, validity } = req.body;

    const FONT_R = path.join(__dirname, 'fonts', 'NotoSans-Regular.ttf');
    const FONT_B = path.join(__dirname, 'fonts', 'NotoSans-Bold.ttf');

    const doc = new PDFDocument({ size: 'A4', margin: 0, info: { Title: `Oferta ${eventName}`, Author: 'flercafe' } });
    doc.registerFont('Regular', FONT_R);
    doc.registerFont('Bold', FONT_B);

    const dateStr = new Date().toLocaleDateString('ro-RO');
    const offerNo = `FC-${Date.now().toString().slice(-6)}`;
    const totalOffer = (items || []).reduce((s, i) => s + (Number(i.totalPrice) || 0), 0);
    const totalCost = Math.round((totalOffer / 2.5) * 100) / 100;
    const profit = Math.round((totalOffer - totalCost) * 100) / 100;
    const EUR_RATE = 5.0;
    const guests = Number(guestCount) || 0;
    const pricePerPersonEUR = guests > 0 ? (totalOffer / guests / EUR_RATE) : 0;
    const pricePerPersonRON = guests > 0 ? (totalOffer / guests) : 0;

    const MARGIN = 50;
    const W = 495;
    const GOLD = '#C8A96E';
    const DARK = '#1a1a1a';
    const GRAY = '#555555';
    const GREEN = '#2d7d46';
    const FOOTER_Y = 775;

    res.setHeader('Content-Type', 'application/pdf');
    const safeFilename = (eventName || 'eveniment')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-zA-Z0-9\-_]/g, '-')
      .replace(/-+/g, '-').replace(/^-|-$/g, '');
    res.setHeader('Content-Disposition', `attachment; filename="Oferta-${safeFilename}-flercafe.pdf"`);
    doc.pipe(res);

    // ═══════════════════ PAGE 1 ═══════════════════

    // Brand header — white background, large black text
    doc.font('Bold').fontSize(30).fillColor(DARK)
       .text('FLERCAFE', MARGIN, 42, { align: 'center', width: W });
    doc.font('Regular').fontSize(9.5).fillColor(GRAY)
       .text('CENTRALIZATOR BUGET & OFERTARE EVENIMENT', MARGIN, 78, { align: 'center', width: W });

    // Double rule under title
    doc.moveTo(MARGIN, 97).lineTo(MARGIN + W, 97).strokeColor(DARK).lineWidth(1.2).stroke();
    doc.moveTo(MARGIN, 100).lineTo(MARGIN + W, 100).strokeColor(GOLD).lineWidth(0.5).stroke();

    // Info block
    const infoY = 116;
    const half = W / 2 - 10;
    const col1X = MARGIN;
    const col2X = MARGIN + W / 2 + 10;

    const drawInfo = (label, value, x, y) => {
      doc.font('Regular').fontSize(7.5).fillColor(GRAY).text(label, x, y);
      doc.font('Bold').fontSize(9).fillColor(DARK).text(value || '—', x, y + 11);
    };
    drawInfo('BENEFICIAR FINAL', clientName || '', col1X, infoY);
    drawInfo('AUDIENTA', `${guestCount || '—'} persoane`, col2X, infoY);
    drawInfo('DATA EVENIMENT', eventDate ? new Date(eventDate).toLocaleDateString('ro-RO') : '—', col1X, infoY + 32);
    drawInfo('MANAGER PROIECT', 'Patrascu Alexandru Metin', col2X, infoY + 32);

    // Light separator
    const sep1Y = infoY + 62;
    doc.moveTo(MARGIN, sep1Y).lineTo(MARGIN + W, sep1Y).strokeColor('#dddddd').lineWidth(0.5).stroke();

    // Section 1 header with gold left border
    const sec1Y = sep1Y + 14;
    doc.rect(MARGIN, sec1Y, 4, 20).fill(GOLD);
    doc.font('Bold').fontSize(9.5).fillColor(DARK)
       .text('1. OFERTARE CLIENT (MARJA 2.5x)', MARGIN + 10, sec1Y + 3);
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text('Preturi finale de ofertare catre client', MARGIN + 10, sec1Y + 16);

    // Products table
    const tY = sec1Y + 36;
    const cols = [
      { label: '#', w: 22, align: 'center' },
      { label: 'PRODUS', w: 185, align: 'left' },
      { label: 'CANT.', w: 50, align: 'center' },
      { label: 'U.M.', w: 42, align: 'left' },
      { label: 'PRET/UM', w: 82, align: 'right' },
      { label: 'TOTAL (LEI)', w: 114, align: 'right' }
    ];
    const colX = [MARGIN];
    cols.forEach((c, i) => colX.push(colX[i] + c.w));

    // Table header
    doc.rect(MARGIN, tY, W, 20).fill(DARK);
    doc.font('Bold').fontSize(7.5).fillColor('#ffffff');
    cols.forEach((c, i) => {
      doc.text(c.label, colX[i] + 3, tY + 6, { width: c.w - 6, align: c.align });
    });

    // Data rows
    let rowY = tY + 20;
    (items || []).forEach((item, idx) => {
      const rh = 18;
      doc.rect(MARGIN, rowY, W, rh).fill(idx % 2 === 0 ? '#f9f7f3' : '#ffffff');
      doc.font('Regular').fontSize(8).fillColor(DARK);
      doc.text(`${idx + 1}`, colX[0] + 3, rowY + 5, { width: cols[0].w - 6, align: 'center' });
      doc.text(item.name || '', colX[1] + 3, rowY + 5, { width: cols[1].w - 6 });
      doc.text(`${item.quantity || ''}`, colX[2] + 3, rowY + 5, { width: cols[2].w - 6, align: 'center' });
      doc.text(item.unit || '', colX[3] + 3, rowY + 5, { width: cols[3].w - 6 });
      doc.text(`${Number(item.unitPrice || 0).toFixed(2)}`, colX[4] + 3, rowY + 5, { width: cols[4].w - 6, align: 'right' });
      doc.font('Bold').fontSize(8).text(`${Number(item.totalPrice || 0).toFixed(2)} lei`, colX[5] + 3, rowY + 5, { width: cols[5].w - 6, align: 'right' });
      rowY += rh;
    });

    // Subtotal row
    doc.rect(MARGIN, rowY, W, 20).fill('#ede9e0');
    doc.font('Bold').fontSize(8).fillColor(DARK);
    doc.text('SUBTOTAL', colX[1] + 3, rowY + 6);
    doc.text(`${totalOffer.toFixed(2)} lei`, colX[5] + 3, rowY + 6, { width: cols[5].w - 6, align: 'right' });
    rowY += 20;

    // Table border outline
    doc.rect(MARGIN, tY, W, rowY - tY).strokeColor('#cccccc').lineWidth(0.5).stroke();

    // Total dark box
    rowY += 12;
    doc.rect(MARGIN, rowY, W, 32).fill(DARK);
    doc.font('Regular').fontSize(9).fillColor('#aaaaaa')
       .text('TOTAL OFERTARE CLIENT:', MARGIN + 12, rowY + 9);
    doc.font('Bold').fontSize(13).fillColor(GOLD)
       .text(`${totalOffer.toFixed(2)} LEI`, MARGIN + 12, rowY + 7, { width: W - 24, align: 'right' });
    rowY += 32;

    // Pret per persoana box (vizibil doar daca stim nr. invitati)
    if (guests > 0) {
      rowY += 8;
      doc.rect(MARGIN, rowY, W, 36).fill('#f5f0e8');
      doc.rect(MARGIN, rowY, 4, 36).fill(GOLD);
      doc.font('Bold').fontSize(18).fillColor(DARK)
         .text(`${pricePerPersonEUR.toFixed(1)} EUR / persoana`, MARGIN + 14, rowY + 5);
      doc.font('Regular').fontSize(9).fillColor(GRAY)
         .text(`(${pricePerPersonRON.toFixed(1)} RON / persoana  ·  ${guests} invitati  ·  curs ${EUR_RATE} RON/EUR)`,
           MARGIN + 14, rowY + 23);
      rowY += 36;
    }

    // Notes (if any)
    if (notes && notes.trim()) {
      rowY += 12;
      doc.font('Bold').fontSize(8).fillColor(DARK).text('Observatii:', MARGIN, rowY);
      rowY += 12;
      doc.font('Regular').fontSize(8).fillColor(GRAY).text(notes, MARGIN, rowY, { width: W });
      rowY += 28;
    }

    // Footer page 1
    doc.moveTo(MARGIN, FOOTER_Y).lineTo(MARGIN + W, FOOTER_Y).dash(3, { space: 3 }).strokeColor('#bbbbbb').lineWidth(0.5).stroke();
    doc.undash();
    doc.font('Bold').fontSize(7.5).fillColor(DARK).text('Intocmit,', MARGIN, FOOTER_Y + 9);
    doc.font('Regular').fontSize(7.5).fillColor(GRAY).text('Patrascu Alexandru Metin', MARGIN, FOOTER_Y + 20);
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text(`Data Documentului: ${dateStr}`, MARGIN, FOOTER_Y + 9, { width: W, align: 'center' });
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text('Pagina 1', MARGIN, FOOTER_Y + 20, { width: W, align: 'right' });

    // ═══════════════════ PAGE 2 ═══════════════════
    doc.addPage();

    // Small brand header
    doc.font('Bold').fontSize(20).fillColor(DARK)
       .text('FLERCAFE', MARGIN, 40, { align: 'center', width: W });
    doc.moveTo(MARGIN, 66).lineTo(MARGIN + W, 66).strokeColor(DARK).lineWidth(0.8).stroke();
    doc.moveTo(MARGIN, 69).lineTo(MARGIN + W, 69).strokeColor(GOLD).lineWidth(0.4).stroke();

    // Section 2 header
    const sec2Y = 88;
    doc.rect(MARGIN, sec2Y, 4, 20).fill(GOLD);
    doc.font('Bold').fontSize(9.5).fillColor(DARK)
       .text('2. PROIECTIE FINANCIARA (UZ INTERN)', MARGIN + 10, sec2Y + 3);
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text('Calculul marjei brute estimate pentru acest eveniment', MARGIN + 10, sec2Y + 16);

    // Financial table
    const ftY = sec2Y + 40;
    const LW = 330;
    const VW = W - LW;

    doc.rect(MARGIN, ftY, W, 20).fill(DARK);
    doc.font('Bold').fontSize(8).fillColor('#ffffff');
    doc.text('INDICATOR FINANCIAR', MARGIN + 8, ftY + 6);
    doc.text('VALOARE (LEI)', MARGIN + LW + 8, ftY + 6, { width: VW - 16, align: 'right' });

    let fy = ftY + 20;

    // Row 1: Total offer
    doc.rect(MARGIN, fy, W, 24).fill('#f9f7f3');
    doc.font('Regular').fontSize(9).fillColor(DARK).text('Ofertare Totala Client', MARGIN + 8, fy + 7);
    doc.font('Bold').fontSize(9).fillColor(DARK)
       .text(`${totalOffer.toFixed(2)} lei`, MARGIN + LW + 8, fy + 7, { width: VW - 16, align: 'right' });
    fy += 24;

    // Row 2: Cost
    doc.rect(MARGIN, fy, W, 24).fill('#ffffff');
    doc.font('Regular').fontSize(9).fillColor(DARK).text('Minus Cost Achizitie Marfa (investitie)', MARGIN + 8, fy + 7);
    doc.font('Bold').fontSize(9).fillColor('#cc3333')
       .text(`- ${totalCost.toFixed(2)} lei`, MARGIN + LW + 8, fy + 7, { width: VW - 16, align: 'right' });
    fy += 24;

    // Row 3: Profit (green highlight)
    doc.rect(MARGIN, fy, W, 28).fill('#e6f4ea');
    doc.font('Bold').fontSize(10).fillColor(GREEN).text('MARJA BRUTA (PROFIT ESTIMAT)', MARGIN + 8, fy + 8);
    doc.font('Bold').fontSize(10).fillColor(GREEN)
       .text(`${profit.toFixed(2)} lei`, MARGIN + LW + 8, fy + 8, { width: VW - 16, align: 'right' });
    fy += 28;

    // Table border
    doc.rect(MARGIN, ftY, W, fy - ftY).strokeColor('#cccccc').lineWidth(0.5).stroke();

    // Note text
    fy += 18;
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text('* Din marja bruta se vor scadea eventualele costuri operationale suplimentare (transport, personal extern, spatii externe etc.)', MARGIN, fy, { width: W });

    // Offer number
    fy += 22;
    doc.font('Regular').fontSize(7.5).fillColor('#aaaaaa')
       .text(`Nr. oferta intern: ${offerNo}`, MARGIN, fy);

    // Footer page 2
    doc.moveTo(MARGIN, FOOTER_Y).lineTo(MARGIN + W, FOOTER_Y).dash(3, { space: 3 }).strokeColor('#bbbbbb').lineWidth(0.5).stroke();
    doc.undash();
    doc.font('Bold').fontSize(7.5).fillColor(DARK).text('Intocmit,', MARGIN, FOOTER_Y + 9);
    doc.font('Regular').fontSize(7.5).fillColor(GRAY).text('Patrascu Alexandru Metin', MARGIN, FOOTER_Y + 20);
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text(`Data Documentului: ${dateStr}`, MARGIN, FOOTER_Y + 9, { width: W, align: 'center' });
    doc.font('Regular').fontSize(7.5).fillColor(GRAY)
       .text('Pagina 2', MARGIN, FOOTER_Y + 20, { width: W, align: 'right' });

    doc.end();
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── HEALTH ─────────────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => res.json({ ok: true, ts: Date.now() }));

// ─── START ──────────────────────────────────────────────────────────────────
process.on('unhandledRejection', (err) => {
  console.error('Unhandled rejection:', err.message || err);
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`flērcafē server pornit pe http://localhost:${PORT}`);

  // Self-ping every 14 min to prevent Render free tier from sleeping
  const selfUrl = process.env.RENDER_EXTERNAL_URL;
  if (selfUrl) {
    const https = require('https');
    setInterval(() => {
      https.get(`${selfUrl}/api/health`, r => r.resume()).on('error', () => {});
    }, 14 * 60 * 1000);
    console.log(`Keep-alive activ → ${selfUrl}/api/health`);
  }
});
