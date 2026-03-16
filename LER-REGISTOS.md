# 🎯 SOLUÇÃO - Ver Tudo do REGISTOS

## ⚡ O Que Mudou

Agora a aplicação **tenta carregar TUDO do Google Sheet REGISTOS primeiro**:

```
1. App abre
2. Tenta ler do Google Sheet REGISTOS
3. Se funcionar → Mostra tudo que lá está
4. Se falhar (CORS) → Mostra questões pré-definidas como fallback
```

---

## 🔧 O Que PRECISA Fazer

### **PASSO 1: Redeployar o AppScript**

Este é o passo **CRÍTICO** para funcionar:

1. Abra o Google Sheet
2. **Extensões > Apps Script**
3. Copie o conteúdo completo de [appscript.gs](../appscript.gs)
4. Cole no Script Editor (apagar tudo primeiro)
5. **Guardar** (Ctrl+S)
6. Clique **"Deploy"** > **"New Deployment"**
   - Tipo: **"Web app"**
   - Executar como: [sua conta Google]
   - Quem tem acesso: **"Qualquer Um"** ⚠️ IMPORTANTE!
7. Copie o **novo URL de deployment**
8. Em [index.html](index.html) linha ~503, confirme/atualize:
   ```javascript
     const APPSCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzFIKdU4JI7tMZ9L1DpcgcnjerWCDle46FGEh157eT2V9urrOgz8J7yZc_HuT2RJuuQ/exec';
   ```

---

### **PASSO 2: Garantir dados em REGISTOS**

Certifique-se que o separador **REGISTOS** tem dados:

| ID | DEPARTAMENTO | Investigação | Foco |
|----|---|---|---|
| Q1 | DCO | Existe registo de pedidos? | Rastreabilidade |
| Q2 | DCO | Os fornecedores estão aprovados? | Qualificação |
| ... | ... | ... | ... |

---

### **PASSO 3: Testar**

1. Recarregue a página (F5)
2. Abra **"Questões"**
3. Na consola (F12), procure por:
   - ✅ Se vir: `"✅ SUCESSO! Questões carregadas do sheet: XX questões"`
   - ❌ Se vir CORS error: URL ainda está desatualizado

---

## 🧪 Debug - Se Não Funcionar

Abra F12 (Console) e execute:

```javascript
// Ver o que está carregado
debugShowData()

// Forçar recarregamento do sheet
debugRecarregar()

// Limpar e recomeçar
debugClear()
```

---

## 📝 O que a Aplicação Faz AGORA

### **Ao Abrir:**
- ✅ Tenta carregar do Google Sheet REGISTOS
- ✅ Se sucesso: Mostra TUDO que lá está
- ✅ Se falha: Usa questões pré-definidas como fallback

### **Na Página "Questões":**
- ✅ Mostra todas as questões por departamento
- ✅ Cada departamento pode ter "+ Adicionar"
- ✅ Novas questões aparecem imediatamente

### **Botão "Recarregar do Sheet":**
- ✅ Força novo carregamento do Google Sheet
- ✅ Atualiza a lista

---

## 🎯 Resumo

```
Google Sheet REGISTOS 
         ↓
    AppScript (com CORS)
         ↓
    Fetch (Browser)
         ↓
    "BASE DE QUESTÕES"
         ↓
    Mostra TUDO
```

**Tudo depende do AppScript estar corretamente redeployado!**

---

**Próximo passo: REDEPLOYE O APPSCRIPT!** 🚀

