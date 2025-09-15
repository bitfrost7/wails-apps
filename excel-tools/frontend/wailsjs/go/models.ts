export namespace backend {
	
	export class KeyWordStatConfig {
	    KwInputDir: string;
	    KwOutputDir: string;
	    StatMode: string;
	    TargetNumber: number;
	    ForwardNumber: number;
	    SelectedColor: string;
	
	    static createFrom(source: any = {}) {
	        return new KeyWordStatConfig(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.KwInputDir = source["KwInputDir"];
	        this.KwOutputDir = source["KwOutputDir"];
	        this.StatMode = source["StatMode"];
	        this.TargetNumber = source["TargetNumber"];
	        this.ForwardNumber = source["ForwardNumber"];
	        this.SelectedColor = source["SelectedColor"];
	    }
	}
	export class WordFreqStatConfig {
	    WfInputDir: string;
	    IntervalNumber: number;
	    SplitChar: string;
	
	    static createFrom(source: any = {}) {
	        return new WordFreqStatConfig(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.WfInputDir = source["WfInputDir"];
	        this.IntervalNumber = source["IntervalNumber"];
	        this.SplitChar = source["SplitChar"];
	    }
	}

}

