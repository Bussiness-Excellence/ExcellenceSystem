# Excellence System

This repository contains the unified codebase for the Excellence CRM project, encompassing both the frontend Vercel web application and the backend automation scripts.

## Folder Structure

- `/webapp`: Contains the React frontend and Vercel serverless functions (`/api`). This folder is deployed directly to Vercel.
- `/scripts`: Contains all the Python and Batch scripts used for backend data extraction, hierarchy management, and uploading to Supabase. These scripts are run locally or on a separate environment.

## Deployment

### Web App
The `webapp` folder is designed to be deployed to Vercel. 
- Build Command: `npm run build`
- Root Directory (in Vercel): `webapp`
- Output Directory: `build` (or depending on your React setup)

### Scripts
The `scripts` folder contains standalone scripts. Please ensure you have the required Python dependencies installed (e.g., `pip install -r requirements.txt` if available) and the necessary environment variables set up.
