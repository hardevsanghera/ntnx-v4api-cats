# Guide: Merging Pagination Code from fixcase to main Branch

This guide explains how to apply ONLY the pagination changes from the `fixcase` branch to the `main` branch.

## 🎯 What Changed - Pagination Feature

The pagination feature allows `list_vms.ps1` to retrieve ALL VMs (not just the first 50-100) by automatically looping through pages.

### Key Changes in list_vms.ps1

**OLD CODE (main branch):**
```powershell
$uri = "$baseUrl/vmm/v4.1/ahv/config/vms?$limit=101"

# --- Make the REST API Call ---
try {
    Write-Host "Calling API at URI: $uri" -ForegroundColor Cyan

    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -SkipCertificateCheck -SkipHttpErrorCheck

    if ($response) {
        Write-Host "Successfully received response." -ForegroundColor Green
        # ... rest of code
    }
}
```

**NEW CODE (fixcase branch with pagination):**
```powershell
# --- Make the REST API Call with Pagination ---
try {
    $pageSize = 100  # Maximum allowed by the API
    $pageNumber = 0
    $allVMs = @()
    $hasMorePages = $true

    while ($hasMorePages) {
        $uri = "$baseUrl/vmm/v4.1/ahv/config/vms?`$limit=$pageSize&`$page=$pageNumber"
        Write-Host "Calling API at URI: $uri (Page $pageNumber)" -ForegroundColor Cyan

        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -SkipCertificateCheck -SkipHttpErrorCheck

        if ($response -and $response.data) {
            $currentPageVMs = $response.data
            $vmCount = ($currentPageVMs | Measure-Object).Count
            Write-Host "Retrieved $vmCount VMs from page $pageNumber" -ForegroundColor Green
            
            $allVMs += $currentPageVMs
            
            # Check if there are more pages
            if ($vmCount -lt $pageSize) {
                $hasMorePages = $false
                Write-Host "Reached last page. Total VMs retrieved: $($allVMs.Count)" -ForegroundColor Yellow
            } else {
                $pageNumber++
            }
        } else {
            Write-Host "No more data returned. Total VMs retrieved: $($allVMs.Count)" -ForegroundColor Yellow
            $hasMorePages = $false
        }
    }

    if ($allVMs.Count -gt 0) {
        Write-Host "Successfully retrieved all $($allVMs.Count) VMs." -ForegroundColor Green

        $vmFile = Join-Path -Path $PWD -ChildPath '\scratch\vm_list.json'

        # Create a combined response object to save
        $combinedResponse = @{
            data = $allVMs
            totalCount = $allVMs.Count
        }

        # Save the combined API response to vm_list.json (overwrite)
        try {
            $combinedResponse | ConvertTo-Json -Depth 10 | Out-File -FilePath $vmFile -Encoding utf8 -Force
            Write-Host "Saved all VMs to $vmFile" -ForegroundColor Yellow
        } catch {
            Write-Error "Failed to write $vmFile :"
            exit 1
        }

        # Read the JSON file (expected structure with entities array)
        try {
            $jsonText = Get-Content -Path $vmFile -Raw -ErrorAction Stop
            # ... rest of code continues as before
```

## 📋 Option 1: Manual Merge (Recommended for Surgical Changes)

### Step 1: Switch to main branch
```powershell
git checkout main
```

### Step 2: Create a backup
```powershell
Copy-Item list_vms.ps1 list_vms.ps1.backup
```

### Step 3: Open list_vms.ps1 and locate the section to replace

Find this section (around line 49-65):
```powershell
$uri = "$baseUrl/vmm/v4.1/ahv/config/vms?$limit=101"

# --- Create the Basic Authentication Header ---
$base64AuthInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($username):$($password)"))
$headers = @{
    Authorization = "Basic $base64AuthInfo"
}

# --- Make the REST API Call ---
try {
    Write-Host "Calling API at URI: $uri" -ForegroundColor Cyan

    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -SkipCertificateCheck -SkipHttpErrorCheck

    if ($response) {
        Write-Host "Successfully received response." -ForegroundColor Green

        $vmFile = Join-Path -Path $PWD -ChildPath '\scratch\vm_list.json'

        # Always save the latest API response to vm_list.json (overwrite)
```

### Step 4: Replace with pagination code

Replace the entire section from after the headers definition through the initial file save with:

```powershell
# --- Make the REST API Call with Pagination ---
try {
    $pageSize = 100  # Maximum allowed by the API
    $pageNumber = 0
    $allVMs = @()
    $hasMorePages = $true

    while ($hasMorePages) {
        $uri = "$baseUrl/vmm/v4.1/ahv/config/vms?`$limit=$pageSize&`$page=$pageNumber"
        Write-Host "Calling API at URI: $uri (Page $pageNumber)" -ForegroundColor Cyan

        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -SkipCertificateCheck -SkipHttpErrorCheck

        if ($response -and $response.data) {
            $currentPageVMs = $response.data
            $vmCount = ($currentPageVMs | Measure-Object).Count
            Write-Host "Retrieved $vmCount VMs from page $pageNumber" -ForegroundColor Green
            
            $allVMs += $currentPageVMs
            
            # Check if there are more pages
            if ($vmCount -lt $pageSize) {
                $hasMorePages = $false
                Write-Host "Reached last page. Total VMs retrieved: $($allVMs.Count)" -ForegroundColor Yellow
            } else {
                $pageNumber++
            }
        } else {
            Write-Host "No more data returned. Total VMs retrieved: $($allVMs.Count)" -ForegroundColor Yellow
            $hasMorePages = $false
        }
    }

    if ($allVMs.Count -gt 0) {
        Write-Host "Successfully retrieved all $($allVMs.Count) VMs." -ForegroundColor Green

        $vmFile = Join-Path -Path $PWD -ChildPath '\scratch\vm_list.json'

        # Create a combined response object to save
        $combinedResponse = @{
            data = $allVMs
            totalCount = $allVMs.Count
        }

        # Save the combined API response to vm_list.json (overwrite)
        try {
            $combinedResponse | ConvertTo-Json -Depth 10 | Out-File -FilePath $vmFile -Encoding utf8 -Force
            Write-Host "Saved all VMs to $vmFile" -ForegroundColor Yellow
        } catch {
            Write-Error "Failed to write $vmFile :"
            exit 1
        }

        # Read the JSON file (expected structure with entities array)
```

**IMPORTANT:** Make sure the backtick escape character before `$limit` and `$page` is preserved: `` `$ ``

### Step 5: Test the changes
```powershell
# Test the script
.\list_vms.ps1

# Verify it retrieves all VMs with pagination messages
```

### Step 6: Commit the changes
```powershell
git add list_vms.ps1
git commit -m "feat: Add pagination support to list_vms.ps1

- Retrieve all VMs by looping through pages of 100
- Escape dollar signs in query parameters ($limit, $page)
- Accumulate all VMs across pages
- Display pagination progress
- Save combined results with total count"

git push origin main
```

## 📋 Option 2: Git Cherry-Pick (For Specific Commits)

If the pagination changes are in a single commit:

### Step 1: Find the commit with pagination changes
```powershell
git checkout fixcase
git log --oneline --grep="pagination" -n 10
# OR
git log --oneline list_vms.ps1 -n 10
```

### Step 2: Note the commit hash
Look for the commit that added pagination (e.g., `abc1234`)

### Step 3: Switch to main and cherry-pick
```powershell
git checkout main
git cherry-pick abc1234
```

### Step 4: Resolve any conflicts
If there are conflicts:
```powershell
# Edit list_vms.ps1 to resolve conflicts
# Keep the pagination code changes
git add list_vms.ps1
git cherry-pick --continue
```

### Step 5: Push to main
```powershell
git push origin main
```

## 📋 Option 3: Create a Patch File

### On fixcase branch:
```powershell
git checkout fixcase

# Create a patch with just the pagination changes
git format-patch main..fixcase -- list_vms.ps1

# This creates a file like: 0001-add-pagination.patch
```

### On main branch:
```powershell
git checkout main

# Apply the patch
git am 0001-add-pagination.patch

# If there are conflicts, resolve them:
git am --show-current-patch
# Edit files to resolve conflicts
git add list_vms.ps1
git am --continue

# Push changes
git push origin main
```

## ✅ Verification Steps

After applying the changes, verify:

1. **Syntax Check:**
   ```powershell
   # Check for PowerShell syntax errors
   Get-Command -Syntax .\list_vms.ps1
   ```

2. **Dry Run Test:**
   ```powershell
   # Run against your test environment
   .\list_vms.ps1
   ```

3. **Verify Output:**
   - Check console output shows pagination: "Page 0", "Page 1", etc.
   - Verify `scratch\vm_list.json` contains all VMs
   - Check that `totalCount` field is present in JSON
   - Verify `scratch\vm_mapping.csv` contains all VM entries

4. **Check for Escaped Dollar Signs:**
   ```powershell
   # Verify the URI contains literal $limit and $page
   Select-String -Path .\list_vms.ps1 -Pattern '`\$limit' -CaseSensitive
   Select-String -Path .\list_vms.ps1 -Pattern '`\$page' -CaseSensitive
   ```

## 🔍 What This Change Does

### Before (Single Request):
- Fetched VMs with `$limit=101` (but API max is 100)
- Only retrieved first 50-100 VMs
- Failed silently if more VMs existed

### After (Pagination):
- Fetches VMs in pages of 100 (API maximum)
- Loops through all pages until fewer than 100 VMs returned
- Accumulates all VMs in `$allVMs` array
- Saves combined results with total count
- Displays progress for each page

### Key Technical Changes:

1. **Escaped Query Parameters:**
   - Changed: `?$limit=100` 
   - To: ``?`$limit=100``
   - Reason: Prevents PowerShell from treating `$limit` as a variable

2. **Pagination Loop:**
   - Tracks `$pageNumber` (starts at 0)
   - Uses `$pageSize = 100` (API maximum)
   - Continues while `$hasMorePages` is true

3. **Stopping Condition:**
   - When `$vmCount < $pageSize`, we've reached the last page
   - OR when `$response.data` is empty/null

4. **Data Accumulation:**
   - `$allVMs += $currentPageVMs` combines results
   - Final save includes all VMs from all pages

## 🐛 Troubleshooting

### Issue: Script still only returns 50 VMs
**Solution:** Check that dollar signs are escaped: `` `$limit `` not `$limit`

### Issue: "Variable $limit not found"
**Solution:** Add backtick before dollar sign: `` `$ ``

### Issue: Merge conflicts
**Solution:** Keep the pagination loop structure from fixcase, but preserve any other changes from main

### Issue: JSON structure different
**Solution:** The new code creates `{ data: [...], totalCount: N }` structure instead of direct response

## 📝 Summary of Lines Changed

**Approximate line numbers in list_vms.ps1:**

- **Lines 49-52**: Remove old single URI construction
- **Lines 53-95**: Add pagination loop with page tracking
- **Lines 96-110**: Update JSON save to use combined response with totalCount

**Total additions:** ~50 lines  
**Total deletions:** ~15 lines  
**Net change:** ~35 additional lines

---

**Created:** November 7, 2025  
**Branch:** fixcase → main  
**Feature:** VM Pagination Support
