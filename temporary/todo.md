# Fill form page

<http://localhost:8080/forms/2224325f0c324380802a9b8efa304009/fill>
ไม่จำเเป็นต้องใช้  plugin 'ตั้งค่า fields' ใช่หรือไม่
ถ้าใช่ให้ซ่อนหรือเอาออกไปได้ user คนที่เข้ามากรอก fields ไม่ควรจะต้องเห็นหรือต้องใช้งาน

# หน้าผลลัพธ์

UI ดูไม่ดีเลย

# Dockerfile for Web

อยากให้เปลี่ยน Nginx เป็น Static web server แทน

นี่เป็นตัวอย่างจากอีก project นึง

```dockerfile
FROM oven/bun:latest AS build

WORKDIR /app

COPY package.json ./
COPY bun.lock ./

RUN bun install

COPY . .
RUN bun docs:build

FROM joseluisq/static-web-server:2
COPY --from=build /app/.vitepress/dist /dist

EXPOSE 3000
CMD [ "--root", "/dist", "--port=3000" ]
```

# config

```toml
[general]
port = 8787
root = "/public"
log-level = "info"
page-fallback = "/public/index.html"    # ✅ SPA fallback

[[headers.rules]]
source = "**"
[headers.rules.headers]
X-Frame-Options = "ALLOWALL"
Content-Security-Policy = "frame-ancestors *"
```

# Dockerfile

```dockerfile
# Stage 1: Build
FROM node:20-alpine AS builder

WORKDIR /app
COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

# Stage 2: Serve
FROM joseluisq/static-web-server:2-alpine

COPY --from=builder /app/dist /public
COPY config.toml /config.toml

EXPOSE 8787

CMD ["--config-file", "/config.toml"]
```

# Onlyoffice

## zen mode

หน้าไหนที่ใช้ Onlyoffice อยากให้มีปุ่มเป็น full screen mode ด้วย
หมายความว่าเวลากดแล้วให้แสดง Onlyoffice ที่เป็น iframe ให้ขยายเต็มจอของ body นะ
ไม่ใช่ไปขยายตัว browser นะ มันนอกขอบเขตของ website แล้ว

เพื่อให้ UX ที่ดีมากขึ้น

## plugin 'ตั้งค่าฟิล์ด'

มัน scroll ไม่ได้ ทำให้ฉันไม่เห็นว่ามีอะไรอยู่ด้านล่าง

# Create user page

อยากให้เพิ่มปุ่ม copy password ชั่วคราวด้วย

# Receipt page

<http://localhost:8080/receipt/1227fdeb-911c-43ca-9f65-52cb8e0b3ce0>

ไม่อยากให้ใช้ json เลย
user ไม่รู้เรื่อง อยากให้แสดงเป็น text ปกติก็ได้ แบบจัดให้สวยงามตาม fields ต่างๆ

หน้า admin result เองก็ไม่อยากให้ใช้ json เหมือนกัน
<http://localhost:8080/admin/results/c33a0068-17b7-41c4-b851-14346e8b548e>

อยากได้แบบหลาย tab

1. view เหมือนตอนที่กรอก docx in onlyoffice แต่ว่า view-only นะ
2. table view คือเอา fiels ต่างๆมาเรียงเป็นตาราง

# AI Agent

ฉันอยากให้ใส่ AI Agent embedded เข้ามาในระบบ มีช่อง chat มี upload pdf file                                                                                                           │
AI Agent ฉันคิดว่าน่าจะลองใช้ Pi ดูก่อน แล้วต่อกับ OpenAI subscription หรือ token ต่างๆ ตามที่ Pi ทําได้เลย อยากใส่ python เพื่อให้ AI Agent เอาไปใช้งานได้ด้วย แต่ยังไม่รู้ว่าจะทําอย่างไร
ถ้าไม่มี login ก็ไม่เปิดให้ใช้ AI                                                                                                                                                     ส่ิงที่ AI Agent จะช่วยก็คือสร้าง form จาก pdf ที่ upload หรือ description ใน chat แล้วสร้างเป็น form ออกมา ฉันคิดว่า Pi น่าจะต้องใช้ python นะแต่ไม่แน่ใจ                                            ใน chat ผู้ใช้งานสามารถ chat กับ Pi เพื่อขอให้สร้าง form และแก้ไขได้เรื่อยๆ ถ้ามี preview pane ได้ด้วยก็จะดีมาก สุดท้ายน่าจะได้ docx มาใช่มะ แล้วให้ user download docx เพื่อเอาไป upload เพื่อสร้าง form
docx ที่ได้จะต้องใส่ static text กับ content-control เพื่อทําเป็น form อะนะ                                                                                                              │

# Another way to filling form

ตอนนี้มีแค่กรอก form ผ่าน onlyoffice ใช่มะ
ฉันอยากได้หน้าที่กรอก form ต่างๆ เป็น form ปกติ เหมือน google form อะ
แต่เอา fields ต่างๆมาจาก form template ที่สร้างไว้
