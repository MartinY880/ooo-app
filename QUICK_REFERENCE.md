# Quick Reference Guide

## 🚀 Getting Started in 5 Steps

### 1️⃣ Configure Azure AD
Edit `src/authConfig.ts`:
```typescript
clientId: "YOUR_CLIENT_ID_HERE"
authority: "https://login.microsoftonline.com/YOUR_TENANT_ID_HERE"
```

### 2️⃣ Configure Power Automate
Edit `src/components/OofForm.tsx`, line ~125:
```typescript
const response = await fetch('YOUR_POWER_AUTOMATE_WEBHOOK_URL', {
```

### 3️⃣ Run the Application
```powershell
npm run dev
```

### 4️⃣ Test the Application
1. Click "Sign in with Microsoft"
2. Enter credentials and consent
3. Fill out the OOF form
4. Test user search (type 2+ characters)
5. Submit the form

### 5️⃣ Deploy to Production
```powershell
# Update redirect URI in authConfig.ts
npm run build
# Deploy the dist/ folder
```

---

## 📁 Key Files to Configure

| File | What to Change | Priority |
|------|---------------|----------|
| `src/authConfig.ts` | Client ID, Tenant ID, Redirect URI | 🔴 Critical |
| `src/components/OofForm.tsx` | Power Automate webhook URL | 🔴 Critical |
| `.env` (optional) | Environment variables | 🟡 Optional |

---

## 🎯 API Permissions Needed in Azure AD

1. **User.Read** - Read signed-in user profile (default)
2. **User.Read.All** - Read all users' profiles (requires admin consent)

Grant admin consent in Azure Portal:
`Azure AD → App registrations → Your App → API permissions → Grant admin consent`

---

## 🔑 Important Code Locations

### Authentication Logic
- **MSAL Setup:** `src/main.tsx` (lines 10-25)
- **Login Button:** `src/App.tsx` (lines 10-60)
- **Config:** `src/authConfig.ts` (all)

### User Search
- **Component:** `src/components/UserSearch.tsx`
- **Debounce Hook:** `src/hooks/useDebounce.ts`
- **Graph API Call:** UserSearch.tsx (lines 55-75)
- **Important:** Look for `ConsistencyLevel: eventual` header

### Form Submission
- **Component:** `src/components/OofForm.tsx`
- **Submit Handler:** Lines 100-150
- **Validation:** Lines 65-95

---

## 🐛 Common Issues & Solutions

### Issue: "Login failed"
✅ **Solution:** Check Client ID and Tenant ID in `authConfig.ts`

### Issue: "User search returns no results"
✅ **Solution:** 
1. Ensure `User.Read.All` permission is granted
2. Grant admin consent in Azure Portal
3. Check `ConsistencyLevel: eventual` header is present

### Issue: "Form submission fails"
✅ **Solution:** 
1. Check Power Automate webhook URL
2. Verify webhook is running
3. Check browser console for errors

### Issue: "Popup blocked"
✅ **Solution:** Allow popups for localhost:5173 in browser

---

## 📊 Data Flow

```
User → Login Button → MSAL Popup → Azure AD
                                      ↓
                                   Token Acquired
                                      ↓
                              App Authenticated
                                      ↓
User Search → Debounce (400ms) → Graph API → Results
                                      ↓
Form Submit → Validation → Power Automate → Success/Error
```

---

## 🎨 Customization Quick Tips

### Change Colors
Edit Tailwind classes in components:
- `bg-blue-600` → `bg-purple-600`
- `text-blue-500` → `text-purple-500`

### Change Debounce Delay
Edit `src/components/UserSearch.tsx`, line ~35:
```typescript
const debouncedSearchTerm = useDebounce(searchTerm, 400); // Change 400
```

### Add Form Fields
Edit `src/components/OofForm.tsx`:
1. Add state: `const [newField, setNewField] = useState('')`
2. Add input in JSX
3. Add to `formData` object in `handleSubmit`

---

## 📞 Support Resources

- **MSAL Docs:** https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview
- **Graph API:** https://docs.microsoft.com/en-us/graph/overview
- **Tailwind CSS:** https://tailwindcss.com/docs
- **React:** https://react.dev

---

## ✅ Pre-Deployment Checklist

- [ ] Azure AD app registered
- [ ] Client ID configured
- [ ] Tenant ID configured
- [ ] API permissions granted
- [ ] Admin consent obtained
- [ ] Power Automate webhook created
- [ ] Webhook URL configured
- [ ] Local testing completed
- [ ] Production redirect URI configured
- [ ] Application built (`npm run build`)
- [ ] HTTPS enabled in production

---

## 🎉 You're Ready!

This application is fully scaffolded and ready to use once you configure:
1. Azure AD credentials
2. Power Automate webhook URL

Run `npm run dev` and start testing! 🚀
