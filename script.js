// 1. Your V2 Contract Address (Confirmed from your Remix screenshot)
const contractAddress = "0xBeFE94981135A6393E0F029C80188b0f582C2a1D";

// 2. The Correct ABI for Register Number (String) Mapping
const abi = [
	{
		"inputs": [
			{ "internalType": "string", "name": "_regNo", "type": "string" },
			{ "internalType": "string", "name": "_name", "type": "string" },
			{ "internalType": "string", "name": "_course", "type": "string" },
			{ "internalType": "string", "name": "_hash", "type": "string" }
		],
		"name": "issueCertificate",
		"outputs": [],
		"stateMutability": "nonpayable",
		"type": "function"
	},
	{
		"inputs": [
			{ "internalType": "string", "name": "_regNo", "type": "string" }
		],
		"name": "getCertificate",
		"outputs": [
			{ "internalType": "string", "name": "", "type": "string" },
			{ "internalType": "string", "name": "", "type": "string" },
			{ "internalType": "string", "name": "", "type": "string" }
		],
		"stateMutability": "view",
		"type": "function"
	},
	{
		"inputs": [],
		"stateMutability": "nonpayable",
		"type": "constructor"
	},
	{
		"inputs": [],
		"name": "universityAdmin",
		"outputs": [
			{ "internalType": "address", "name": "", "type": "address" }
		],
		"stateMutability": "view",
		"type": "function"
	}
];

async function verifyCertificate() {
    const regNo = document.getElementById("regInput").value;
    const resultDiv = document.getElementById("result");
    
    // Basic validation to ensure the user typed something
    if (!regNo) {
        alert("Please enter a Register Number.");
        return;
    }

    try {
        // Connect to MetaMask/Ethereum Provider
        const provider = new ethers.providers.Web3Provider(window.ethereum);
        
        // Create the contract instance using our V2 Address and ABI
        const contract = new ethers.Contract(contractAddress, abi, provider);

        // Update UI to show we are searching
        resultDiv.style.display = "block";
        resultDiv.innerHTML = "Searching Decentralized Ledger...";
        resultDiv.className = ""; 

        // CRITICAL: We call 'getCertificate' because your V2 code uses this function name
        // We pass the regNo as a STRING
        const data = await contract.getCertificate(regNo);
        
        // If the call succeeds, display the data
        resultDiv.className = "success";
        resultDiv.innerHTML = `
            <div style="text-align: left;">
                <strong style="color: #3fb950;">✅ AUTHENTIC RECORD FOUND</strong><br><br>
                <strong>Student Name:</strong> ${data[0]}<br>
                <strong>Course:</strong> ${data[1]}<br>
                <strong>File Hash:</strong> ${data[2]}
            </div>
        `;
    } catch (error) {
        // If the Register Number doesn't exist, the contract 'reverts' and triggers this
        resultDiv.style.display = "block";
        resultDiv.className = "error";
        resultDiv.innerHTML = "❌ Verification Failed: No record exists for this Register Number.";
        console.error("Blockchain Error:", error);
    }
}