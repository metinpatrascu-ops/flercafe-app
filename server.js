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

// ─── MONGOOSE ──────────────────────────────────────────────────────────────
mongoose.connect(process.env.MONGODB_URI || 'mongodb://localhost:27017/flercafe');

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
  notes: String
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
    const event = await Event.create(req.body);
    res.json(event);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.put('/api/events/:id', verifyToken, async (req, res) => {
  const event = await Event.findByIdAndUpdate(req.params.id, req.body, { new: true });
  res.json(event);
});

// Calculator HoReCa bazat pe reguli — funcționează fără OpenAI
function ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems }) {
  const n = guestCount;
  const h = durationHours || 4;
  const isCorporate = ['corporate', 'lansare'].includes(eventType);
  const isSummer = season === 'vara';
  const MARGIN = 1.15;

  const waterMult = (isCorporate ? 1.2 : 1) * (isSummer ? 1.25 : 1);
  const iceMult = isSummer ? 1.3 : 1;
  const durationMult = h / 4;

  // Detectează meniu cerut din menuItems
  const menuText = (menuItems || []).map(m => `${m.name} ${m.quantity} ${m.unit}`).join(' ').toLowerCase();
  const hasProsecco = menuItems && menuItems.some(m => /prosecco|sampanie|vin spumant/i.test(m.name));
  const proseccoCount = hasProsecco ? (menuItems.find(m => /prosecco|sampanie/i.test(m.name))?.quantity || 0) : 0;

  const rec = [];

  // Apă plată
  const waterStillL = n * 0.75 * waterMult * durationMult * MARGIN;
  rec.push({
    productName: 'Apă Plată 0.75L',
    category: 'apa',
    estimatedConsumptionL: +(waterStillL).toFixed(1),
    recommendedQuantity: Math.ceil(waterStillL / (0.75 * 12)),
    recommendedUnit: 'baxuri (12 sticle × 0.75L)',
    reason: `0.75L/pers × ${n} inv. × factori ajustare ${isCorporate ? 'corporate' : ''} ${isSummer ? 'vară' : ''} + 15% marjă`,
    warningMessage: null
  });

  // Apă minerală
  const waterSparkL = n * 0.5 * waterMult * durationMult * MARGIN;
  rec.push({
    productName: 'Apă Minerală 0.75L',
    category: 'apa',
    estimatedConsumptionL: +(waterSparkL).toFixed(1),
    recommendedQuantity: Math.ceil(waterSparkL / (0.75 * 12)),
    recommendedUnit: 'baxuri (12 sticle × 0.75L)',
    reason: `0.5L/pers × ${n} inv. + ajustări ${isSummer ? 'vară' : ''} + 15% marjă`,
    warningMessage: null
  });

  // Soft drinks (Cola etc.)
  const softL = n * 0.5 * durationMult * MARGIN;
  rec.push({
    productName: 'Cola / Soft Drinks 0.33L',
    category: 'soft',
    estimatedConsumptionL: +(softL).toFixed(1),
    recommendedQuantity: Math.ceil(softL / (0.33 * 24)),
    recommendedUnit: 'baxuri (24 sticle × 0.33L)',
    reason: `~1.5 sticle/pers × ${n} invitați + 15% marjă`,
    warningMessage: null
  });

  // Socată / crafting
  const craftL = n * 0.35 * durationMult * MARGIN;
  rec.push({
    productName: 'Socată / Băutură Crafting',
    category: 'sirop',
    estimatedConsumptionL: +(craftL).toFixed(1),
    recommendedQuantity: Math.ceil(craftL * 1000 / 300),
    recommendedUnit: 'sticle / porții (300ml/porție)',
    reason: `0.35L/pers × ${n} invitați + 15% marjă`,
    warningMessage: null
  });

  // Cocktailuri
  const cocktailPerPerson = isCorporate ? 1.8 : 2.5;
  const cocktailsTotal = Math.ceil(n * cocktailPerPerson * durationMult * MARGIN);
  rec.push({
    productName: 'Cocktailuri (total porții)',
    category: 'altele',
    estimatedConsumptionL: null,
    recommendedQuantity: cocktailsTotal,
    recommendedUnit: 'porții',
    reason: `${cocktailPerPerson} cocktailuri/pers × ${n} inv. ${isCorporate ? '(redus pt. corporate)' : ''} + 15%`,
    warningMessage: isCorporate ? 'Eveniment corporate — consumul de alcool e mai mic. Compensează cu apă și soft.' : null
  });

  // Prosecco (dacă apare în meniu)
  if (proseccoCount > 0) {
    const pahare = proseccoCount * 6;
    const warn = pahare < n ? `⚠️ ${proseccoCount} sticle = ${pahare} pahare. Pentru ${n} invitați nu ajunge! Recomandăm minim ${Math.ceil(n / 6)} sticle.` : null;
    rec.push({
      productName: 'Prosecco / Vin Spumant (welcome)',
      category: 'alcool',
      estimatedConsumptionL: null,
      recommendedQuantity: warn ? Math.ceil(n / 6) : proseccoCount,
      recommendedUnit: `sticle (1 sticlă = 6 pahare)`,
      reason: `Welcome drink pentru ${n} invitați`,
      warningMessage: warn
    });
  }

  // Gheață
  const iceKg = Math.ceil((n * 0.4 + (proseccoCount * 0.3)) * iceMult * MARGIN);
  rec.push({
    productName: 'Gheață',
    category: 'consumabile',
    estimatedConsumptionL: null,
    recommendedQuantity: Math.ceil(iceKg / 5),
    recommendedUnit: `saci de 5kg (total ~${iceKg}kg)`,
    reason: `400g/pers pentru bar + 300g per sticlă Prosecco + 15% marjă ${isSummer ? '+ 30% vară' : ''}`,
    warningMessage: isSummer ? 'Sezon cald — gheața se topește rapid. Asigură-te că ai congelator/ladă frigorifică.' : null
  });

  const warnings = [];
  if (isCorporate) warnings.push('Eveniment corporate/lansare: invitații vin cu mașina. Crește apa și soft-urile, reduci alcoolul.');
  if (isSummer) warnings.push('Temperaturi ridicate de vară: hidratarea este critică. Asigură-te că produsele sunt răcite înainte de eveniment.');
  if (proseccoCount > 0 && proseccoCount * 6 < n) warnings.push(`Ai cerut ${proseccoCount} sticle Prosecco welcome, dar ai ${n} invitați. Nu vor ajunge!`);

  return {
    summary: `Eveniment ${eventType} cu ${n} invitați pe durata de ${h} ore, sezon ${season}. ${isCorporate ? 'Profil corporate — consum moderat de alcool, mare de apă și soft.' : ''} Toate cantitățile includ o marjă de siguranță de +15%.`,
    warnings,
    recommendations: rec,
    usedAI: false
  };
}

app.post('/api/events/analyze', verifyToken, async (req, res) => {
  try {
    const { guestCount, durationHours, eventType, season, briefText, menuItems } = req.body;
    if (!guestCount || guestCount < 1) return res.status(400).json({ error: 'Număr invitați invalid' });

    // Fără cheie OpenAI → calculator pe bază de reguli
    if (!HAS_OPENAI) {
      const analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems });
      return res.json({ analysis, usedAI: false });
    }

    const menuText = (menuItems || []).map(m => `- ${m.name}: ${m.quantity} ${m.unit}`).join('\n');
    const prompt = `Ești un consultant expert în baruri de evenimente și HoReCa din România.
Analizează acest brief de eveniment și calculează EXACT ce cantități trebuie achiziționate.

BRIEF: ${briefText || 'Nu a fost furnizat.'}
Invitați: ${guestCount}, Durată: ${durationHours}h, Tip: ${eventType}, Sezon: ${season}
Meniu client:\n${menuText || '(nespecificat)'}

REGULI HoReCa:
1. Apă plată: 0.75L/pers/4ore (+25% vara, +20% corporate)
2. Apă minerală: 0.5L/pers/4ore (aceleași majorări)
3. Soft drinks: 0.5L/pers
4. Băuturi crafting/socată: 0.35L/pers
5. Cocktailuri: 2.5/pers (−30% corporate)
6. Prosecco: 1 sticlă = 6 pahare 125ml
7. Gheață: 400g/pers bar + 300g/sticlă prosecco (+30% vara)
8. Adaugă ÎNTOTDEAUNA +15% marjă de siguranță

RETURNEAZĂ STRICT JSON:
{"summary":"...","warnings":["..."],"recommendations":[{"productName":"...","category":"apa|soft|alcool|sirop|cafea|consumabile","estimatedConsumptionL":0,"recommendedQuantity":0,"recommendedUnit":"...","reason":"...","warningMessage":null}],"usedAI":true}`;

    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 2500,
      temperature: 0.2
    });

    let aiText = response.choices[0].message.content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    let analysis;
    try { analysis = JSON.parse(aiText); }
    catch { analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems }); }

    res.json({ analysis, usedAI: true });
  } catch (e) {
    // Fallback la reguli dacă AI-ul pică
    try {
      const { guestCount, durationHours, eventType, season, menuItems } = req.body;
      const analysis = ruleBasedAnalysis({ guestCount, durationHours, eventType, season, menuItems });
      res.json({ analysis, usedAI: false });
    } catch (e2) {
      res.status(500).json({ error: e.message });
    }
  }
});

// ─── STOCK ──────────────────────────────────────────────────────────────────
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

// ─── START ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`flērcafē server pornit pe http://localhost:${PORT}`));
