(function() {
    let shadowRoot;
	
	let template = document.createElement("template");
	template.innerHTML = `<style>
					* {
				      margin: 0;
				      padding: 0;
				  }
				  body {
				      background: #fff;
				      font-family: 'Open-Sans',sans-serif;
				  }
				</style>`
				
	const MSTeamslib = "https://statics.teams.cdn.office.net/sdk/v1.6.0/js/MicrosoftTeams.min.js";	


	function loadScript(src) {
	  return new Promise(function(resolve, reject) {
		let script = document.createElement('script');
		console.log("¦¦¦¦¦¦¦¦¦¦¦¦ Load script ¦¦¦¦¦¦¦¦¦¦");
		console.log(src);	    
		console.log("¦¦¦¦¦¦¦¦¦¦¦¦ Load script ¦¦¦¦¦¦¦¦¦¦");	    
		script.src = src;

		script.onload = () => {console.log("Load: " + src); resolve(script);}
		script.onerror = () => reject(new Error(`Script load error for ${src}`));

		shadowRoot.appendChild(script)
	  });
	}	
				
	};
	
	
	
	   class MSTeams extends HTMLElement {
        constructor() {
	    console.log("-------------------------------------------------");	
            console.log("constructor");
	    console.log("-------------------------------------------------");	
            super();
            shadowRoot = this.attachShadow({
                mode: "open"
            });

            shadowRoot.appendChild(template.content.cloneNode(true));

            this._firstConnection = 0;

            this.addEventListener("click", event => {
                console.log('click');
                var event = new Event("onClick");
                this.dispatchEvent(event);

            });
            this._props = {};
        }

        //Fired when the widget is added to the html DOM of the page
		connectedCallback() {
            console.log("connectedCallback");
        }

		//Fired when the widget is removed from the html DOM of the page (e.g. by hide)
		disconnectedCallback() {
			console.log("disconnectedCallback");
        }

		//When the custom widget is updated, the Custom Widget SDK framework executes this function first
        onCustomWidgetBeforeUpdate(changedProperties) {
            console.log("onCustomWidgetBeforeUpdate");
            this._props = {
                ...this._props,
                ...changedProperties
            };
        }
		
		if (this._firstConnection === 0) {
		async function LoadLibs() {
			try {
				await loadScript(MSTeamslib);	
			} catch (e) {
				alert(e);
			} finally {	
				that._firstConnection = 1;	
			}
		}
		LoadLibs();
		}
		
		else {
		
		microsoftTeams.initialize();
		}
		
				//When the custom widget is removed from the canvas or the analytic application is closed
        	onCustomWidgetDestroy() {
		console.log("onCustomWidgetDestroy");
        }
		

	   customElements.define("com-karamba-msteams", MSTeams);
	
	
})();	
