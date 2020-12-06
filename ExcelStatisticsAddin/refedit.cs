using System;
using System.ComponentModel;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel; 
using System.Reflection;

namespace VS.NET_RefeditControl
{
	/// <summary>
	/// Event handler for the CellChange Event
	/// </summary>
	/// <param name="sender"></param>
	/// <param name="e"></param>
	public delegate void refeditCellChangeEventHandler(object sender, EventArgs e);

	/// <summary>
	/// Event handler for Restore (form)  Event
	/// </summary>
	/// <param name="sender"></param>
	/// <param name="e"></param>
	public delegate void refeditRestoreEventHandler(object sender, EventArgs e);

	/// <summary>
	/// Excel Refedit Control.
	/// </summary>
	[ToolboxBitmap(typeof(refedit)), DefaultProperty("Relativity"), DefaultEvent("CellChanged")]
	public class refedit : UserControl
	{

		/// <summary>
		/// The event to raise when the cell reference changes
		/// </summary>
		[Description("Occurs when the cell reference changes")]
		public event refeditCellChangeEventHandler CellChanged;
		
		/// <summary>
		/// The event to raise when the refedit control restores
		/// </summary>
		public event refeditRestoreEventHandler Restore;
		
		/// <summary>
		/// Clicks the refedit button
		/// </summary>
		public void DoClick()
		{
			button1_Click(button1, null);
		}
		
		/// <summary>
		/// Sets the refedit control as the active control
		/// </summary>
		/// <param name="state"></param>
		public void isActive(bool state)
		{
			timer1.Enabled = state;
			_Entered = state;
		}
		/// <summary>
		/// Raises the Cell Changed event
		/// </summary>
		/// <param name="e">EventArgs</param>
		protected virtual void OnCellChanged(EventArgs e)
		{
			if (CellChanged != null) 
			{
				// Invokes the delegates. 
				CellChanged(this, e);
			}
		}
		/// <summary>
		/// Raises the restore event
		/// </summary>
		/// <param name="e">EventArgs</param>
		protected virtual void OnRestore(EventArgs e)
		{
			if (Restore != null) 
			{
				// Invokes the delegates. 
				Restore(this, e);
			}
		}

		#region Variables/constants

		private Timer timer1;
		private Button button1;
		private TextBox textRange;
		private IContainer components;
		private Excel.Application oExcel;
		private const Excel.XlReferenceStyle A1 = Excel.XlReferenceStyle.xlA1;
		private string _CollapsedFormCaption = "Select Range";
		private string parentCaption;
		private string _ActiveSheetName = "";
		private string thisBookName = "";
		private string thisSheetName = "";
		private string _DisplayResultSeparator = " ";
		private ResultSeparator _ResultSeparator = ResultSeparator.Space;
		private HorizontalAlignment _textAlign = HorizontalAlignment.Left;
		private bool _AllowMultipleCells = true;
		private bool _AllowCollapsedResize;
		private bool _AllowTextEntry = true;
		private bool _showEditField = true;
		private bool oldshowEditField = true;
		private bool excelMoveCursorOnEnter = true;
		private int _bigSize = 150;
		private bool isCollapsed;
		private bool _Entered;
		private bool RowAbsolute;
		private bool ColAbsolute;
		private int _adjustHeight;
		private bool topMostForm;
		private Point thisLocation;
		private Size thisSize;
		private Size thisFormSize;
		private Control parentControl;
		private Form thisForm;
		private Form MDIForm;
		private MainMenu thisFormMenu;
		private MainMenu MDIFormMenu;
		private readonly Logger _xlLogger = new Logger();
		private AnchorStyles anchorControl;
		private DockStyle dockStyle;
		private FormBorderStyle borderStyle;

		private ReferenceStyle _Relativity = ReferenceStyle.Absolute;
		private CollapseStyle _CollapseStyle = CollapseStyle.CollapseFormAndFitCellSelector;
		private DisplayStyle _DisplayStyle = DisplayStyle.AddressRange;

		private Hashtable visibles;
		private ArrayList parents;
		private ImageList imageList1;
		private Hashtable MDIvisibles;
		/// <summary>
		/// required calls
		/// </summary>
		/// <param name="nVirtKey">na</param>
		/// <returns>na</returns>
		[DllImport("user32.dll")]	protected static extern int GetKeyState(int nVirtKey);
		/// <summary>
		/// required calls
		/// </summary>
		/// <param name="pbKeyState">na</param>
		/// <returns>na</returns>
		[DllImport("user32.dll")]	protected static extern int GetKeyboardState(byte[] pbKeyState);
		/// <summary>
		/// required calls
		/// </summary>
		/// <param name="lppbKeyState">na</param>
		/// <returns>na</returns>
		[DllImport("user32.dll")]	protected static extern int SetKeyboardState(byte[] lppbKeyState);

		private const int ENTER_KEY = 13;
		#endregion
		
		#region Constructor/Overrides
		/// <summary>
		/// Creates an instance of the Control
		/// </summary>
		public refedit()
		{
			// This call is required by the Windows.Forms Form Designer.
			Application.VisualStyleState = VisualStyleState.ClientAndNonClientAreasEnabled;
			Application.EnableVisualStyles();
			
			InitializeComponent();

		}
		
		
		public refedit(Excel.Application excel)
		{
			// This call is required by the Windows.Forms Form Designer.
			Application.VisualStyleState = VisualStyleState.ClientAndNonClientAreasEnabled;
			Application.EnableVisualStyles();
			
			InitializeComponent();
			oExcel = excel;
		}

		private void refedit_Load(object sender, EventArgs e)
		{
			try
			{ 
				oldshowEditField = ShowEditField;
				thisForm = ParentForm;
				if (thisForm != null)
				{
					topMostForm = thisForm.TopMost;
					thisFormMenu = thisForm.Menu;
				}
				parentControl = Parent;

				//thisForm.LostFocus+=new EventHandler(ParentForm_LostFocus);

				if (thisForm!=null && thisForm.MdiParent != null) 
				{
					MDIForm = thisForm.MdiParent; 
					MDIFormMenu = MDIForm.Menu; 
					topMostForm = MDIForm.TopMost;
					parentCaption = MDIForm.Text;
				}
				if(_Excel!=null)
				{
					var mSheet = ((Excel.Worksheet)(_Excel.ActiveSheet)); 
					_ActiveSheetName = mSheet.Name;
					excelMoveCursorOnEnter = _Excel.MoveAfterReturn;
					_Excel.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(excelEvents_SheetSelectionChange);
				}
				textRange.TextAlign = _textAlign;
			} 
			catch (Exception ex)
			{ 
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
			killEnterKey();
		}

		/// <summary>
		/// Clean up any resources
		/// </summary>
		/// <param name="disposing">true</param>
		protected override void Dispose( bool disposing )
		{
			try
			{
				try
				{
					//Application.MoveAfterReturn = False
					//Application.MoveAfterReturn = True
					if (oExcel != null)
					{
						_Excel.MoveAfterReturn = excelMoveCursorOnEnter;
						oExcel = null;
					}
				}
				catch { }
				if( disposing )
				{
					if(components != null)
					{
						components.Dispose();
					}
				}
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// Occurs when the control is resized
		/// </summary>
		/// <param name="e"></param>
		protected override void OnResize(EventArgs e)
		{
			base.OnResize (e);
			
			//Make the button square
			button1.Height = textRange.Height;
			button1.Width = button1.Height;
			
			//make the height of the control match the textbox
			Height = textRange.Height;
			
			//Now, hide the text box, if property is set
			if(!_showEditField)
			{
				Size = button1.Size;
			}

			Invalidate();
		}

		#endregion

		#region Properties
		/// <summary>
		/// A reference to the Excel Application
		/// </summary>
		[Browsable(false)] public Excel.Application _Excel
		{
			get
			{
				try
				{
					if(oExcel!=null)
					{
						return oExcel;
					}
				}
				catch (Exception ex)
				{
					_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
				}
				return null;
			}
			set
			{
				try
				{
					oExcel = value;
				}
				catch (Exception ex)
				{
					_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
				}
			}
		}
		/// <summary>
		/// Returns the Collapsed state of the control
		/// </summary>
		[Browsable(false)] public bool collapsedButtonState
		{
			get
			{
				return isCollapsed;
			}
		}

		/// <summary>
		/// Indicates how the text should be aligned
		/// </summary>
		[Category("Appearance"), Description("Indicates how the text should be aligned"), DefaultValue(HorizontalAlignment.Left)] public HorizontalAlignment TextAlign
		{
			get
			{
				return _textAlign;
			}
			set
			{
				_textAlign = value;
			}
		}
		/// <summary>
		/// The Row/Column relatively of the displayed range
		/// </summary>
		[Category("Refedit"), Description("The Row/Column relatively of the displayed range"), DefaultValue(ReferenceStyle.Absolute)] public ReferenceStyle Relativity
		{
			get
			{
				return _Relativity;
			}
			set
			{
				_Relativity = value;
				RowAbsolute = (_Relativity == ReferenceStyle.Absolute | _Relativity == ReferenceStyle.ColumnRelative);
				ColAbsolute = (_Relativity == ReferenceStyle.Absolute | _Relativity == ReferenceStyle.RowRelative);
			}
		}
		/// <summary>
		/// The Text displayed in the control
		/// </summary>
		[Category("Refedit"), Description("The Text displayed in the control"), DefaultValue("")] public override string Text
		{
			get
			{
				return textRange.Text;
			}
			set
			{
				textRange.Text = value;
			}
		}
		/// <summary>
		/// The behaviour of the parent form when the selector button is clicked
		/// </summary>
		[Category("Refedit"), Description("The behaviour of the parent form when the selector button is clicked"), DefaultValue(CollapseStyle.CollapseFormAndFitCellSelector)]
		public CollapseStyle CollapseFormStyle
		{
			get
			{
				return _CollapseStyle;
			}
			set
			{
				_CollapseStyle = value;
			}
		}
		/// <summary>
		/// The separator to use when combining multiple results in the Cell Selector (ignored unless DisplayStyle is set to ConcatenateRangeValues
		/// </summary>
		[Category("Refedit"), Description("The separator to use when combining multiple results in the Cell Selector (ignored unless DisplayStyle is set to ConcatenateRangeValues"), DefaultValue(ResultSeparator.Space)]
		public ResultSeparator DisplayResultSeparator
		{
			get
			{
				switch(_ResultSeparator)
				{
					case ResultSeparator.None:
						_DisplayResultSeparator="";
						break;
					case ResultSeparator.Colon:
						_DisplayResultSeparator=":";
						break;
					case ResultSeparator.Pipe:
						_DisplayResultSeparator="|";
						break;
					case ResultSeparator.SemiColon:
						_DisplayResultSeparator=";";
						break;
					case ResultSeparator.Space:
						_DisplayResultSeparator=" ";
						break;
					case ResultSeparator.Comma:
						_DisplayResultSeparator=",";
						break;
				}
				return _ResultSeparator;
			}
			set
			{
				_ResultSeparator = value;
				switch(_ResultSeparator)
				{
					case ResultSeparator.None:
						_DisplayResultSeparator="";
						break;
					case ResultSeparator.Colon:
						_DisplayResultSeparator=":";
						break;
					case ResultSeparator.Pipe:
						_DisplayResultSeparator="|";
						break;
					case ResultSeparator.SemiColon:
						_DisplayResultSeparator=";";
						break;
					case ResultSeparator.Space:
						_DisplayResultSeparator=" ";
						break;
					case ResultSeparator.Comma:
						_DisplayResultSeparator=",";
						break;
				}
			}
		}
		/// <summary>
		/// The type of results to display in the Cell Selector when a cell or range of cells is selected
		/// </summary>
		[Category("Refedit"), Description("The type of results to display in the Cell Selector when a cell or range of cells is selected"), DefaultValue(DisplayStyle.AddressRange)]
		public DisplayStyle DisplayResultStyle
		{
			get
			{
				return _DisplayStyle;
			}
			set
			{
				_DisplayStyle = value;
			}
		}
		/// <summary>
		/// The caption to display on the form while it is in a collapsed state
		/// </summary>
		[Category("Refedit"), Description("The caption to display on the form while it is in a collapsed state"), DefaultValue("Select Range")]
		public string CollapsedFormCaption
		{
			get
			{
				return _CollapsedFormCaption;
			}
			set
			{
				_CollapsedFormCaption = value;
			}
		}
		/// <summary>
		/// Determines whether the control's parent form can be resized when it is in Collapsed mode
		/// </summary>
		[Category("Refedit"), Description("Determines whether the control's parent form can be resized when it is in Collapsed mode"), DefaultValue(true)]
		public bool AllowCollapsedFormResize
		{
			get
			{
				return _AllowCollapsedResize;
			}
			set
			{
				_AllowCollapsedResize = value;
			}
		}
		/// <summary>
		/// Determines whether single cell or multiple cell selection is allowed
		/// </summary>
		[Category("Refedit"), Description("Determines whether single cell or multiple cell selection is allowed"), DefaultValue(true)]
		public bool AllowMultipleCells
		{
			get
			{
				return _AllowMultipleCells;
			}
			set
			{
				_AllowMultipleCells = value;
			}
		}
		/// <summary>
		/// Determines whether typing into this control is Allowed (If False, a valid entry can only be made by selecting cell reference(s) in a worksheet
		/// </summary>
		[Category("Refedit"), Description("Determines whether typing into this control is Allowed (If False, a valid entry can only be made by selecting cell reference(s) in a worksheet"), DefaultValue(true)]
		public bool AllowTextEntry
		{
			get
			{
				return _AllowTextEntry;
			}
			set
			{
				_AllowTextEntry = value;
			}
		}
		/// <summary>
		/// Display the Text Edit control. If False, only the selector button is displayed
		/// </summary>
		[Category("Refedit"), Description("Display the Text Edit control. If False, only the selector button is displayed"), DefaultValue(true)]
		public bool ShowEditField
		{
			get
			{
				return _showEditField;
			}
			set
			{
				_showEditField = value;
				if(!_showEditField)
				{
					_bigSize = Width;
					Size = button1.Size;
				}
				else
				{
					Width = _bigSize;
				}
				Invalidate(true);
			}
		}
		/// <summary>
		/// A unit of Measure that can be used to correct the Form Height when in it's collapsed state.
		/// </summary>
		[Category("Refedit"), Description("A unit of Measure that can be used to correct the Form Height when in it's collapsed state."), DefaultValue(0)]
		public int CollapsedHeightAdjustment
		{
			get
			{
				return _adjustHeight;
			}
			set
			{
				_adjustHeight = value;
			}
		}

		#endregion

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(refedit));
			timer1 = new System.Windows.Forms.Timer(components);
			button1 = new System.Windows.Forms.Button();
			imageList1 = new System.Windows.Forms.ImageList(components);
			textRange = new System.Windows.Forms.TextBox();
			SuspendLayout();
			// 
			// timer1
			// 
			timer1.Tick += new System.EventHandler(timer1_Tick);
			// 
			// button1
			// 
			button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			button1.Cursor = System.Windows.Forms.Cursors.Hand;
			button1.ImageIndex = 0;
			button1.ImageList = imageList1;
			button1.Location = new System.Drawing.Point(152, 0);
			button1.Name = "button1";
			button1.Size = new System.Drawing.Size(20, 20);
			button1.TabIndex = 0;
			button1.TextAlign = System.Drawing.ContentAlignment.TopLeft;
			button1.Click += new System.EventHandler(button1_Click);
			// 
			// imageList1
			// 
			imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			imageList1.TransparentColor = System.Drawing.Color.Transparent;
			imageList1.Images.SetKeyName(0, "refedit.bmp");
			imageList1.Images.SetKeyName(1, "refeditUp.bmp");
			// 
			// textRange
			// 
			textRange.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
			                                                               | System.Windows.Forms.AnchorStyles.Left)
			                                                              | System.Windows.Forms.AnchorStyles.Right)));
			textRange.Location = new System.Drawing.Point(0, 0);
			textRange.Name = "textRange";
			textRange.Size = new System.Drawing.Size(152, 20);
			textRange.TabIndex = 1;
			textRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			textRange.WordWrap = false;
			textRange.Enter += new System.EventHandler(refedit_Enter);
			textRange.Leave += new System.EventHandler(refedit_Leave);
			textRange.KeyUp += new System.Windows.Forms.KeyEventHandler(refedit_KeyUp);
			textRange.TextChanged += new System.EventHandler(textRange_TextChanged);
			textRange.KeyDown += new System.Windows.Forms.KeyEventHandler(refedit_KeyDown);
			// 
			// refedit
			// 
			AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
			Controls.Add(textRange);
			Controls.Add(button1);
			DoubleBuffered = true;
			Name = "refedit";
			Size = new System.Drawing.Size(172, 20);
			Enter += new System.EventHandler(refedit_Enter);
			Load += new System.EventHandler(refedit_Load);
			Resize += new System.EventHandler(refedit_Resize);
			KeyUp += new System.Windows.Forms.KeyEventHandler(refedit_KeyUp);
			KeyDown += new System.Windows.Forms.KeyEventHandler(refedit_KeyDown);
			ResumeLayout(false);
			PerformLayout();
		}
		#endregion

		#region Events
		private void setChildren(Control o, bool setAs)
		{
			foreach (Control x in o.Controls)
			{
				if (x is refedit)
				{
					if (x != this)
					{
						((refedit)x).isActive(false);
					}
				}
				else if (x.HasChildren)
				{
					setChildren(x, false);
				}
			}
		}
		private void setThisControlActive()
		{
			isActive(true);

			if (ParentForm != null)
				foreach(Control x in ParentForm.Controls)
				{
					if (x is refedit)
					{
						if (x != this)
						{
							((refedit)x).isActive(false);
						}
					}
					else if (x.HasChildren)
					{
						setChildren(x, false);
					}

				}
		}

		private void refedit_Enter(object sender, EventArgs e)
		{
			if(DesignMode) return;
			
			if(!Visible || !Enabled) return;
			
			_Entered = true;

			setThisControlActive();
			
			try
			{
				((Excel.Worksheet)_Excel.ActiveSheet).get_Range(Text, Type.Missing).Select();
			}
			catch { }
			timer1.Enabled = true;
			timer1_Tick(null, null);
		}
		private void refedit_Leave(object sender, EventArgs e)
		{
			if(DesignMode) return;

			if(!Visible || !Enabled) return;
			
			_Entered = false;
			timer1.Enabled = false;
		}
		private void refedit_Resize(object sender, EventArgs e)
		{
			try
			{
				if(DesignMode) return;

				if(!Visible || !Enabled) return;
			
				textRange.Location = new Point(0, 0); 
				textRange.Size = new Size(Width - button1.Width, Height); 
				button1.Location = new Point(textRange.Width, 0); 
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}

		private void refedit_KeyDown(object sender, KeyEventArgs e)
		{
			try
			{
				if(DesignMode) return;

				if(!Visible || !Enabled) return;
			
				timer1.Enabled = false;
				if ((e.KeyCode == Keys.Enter) && isCollapsed)
				{
					button1_Click(null, null);
					return;
				}
				if (e.KeyCode == Keys.F4) 
				{ 
					if (_Relativity == ReferenceStyle.Absolute) 
					{ 
						Relativity = ReferenceStyle.RowRelative; 
					} 
					else if (_Relativity == ReferenceStyle.RowRelative) 
					{ 
						Relativity = ReferenceStyle.ColumnRelative; 
					} 
					else if (_Relativity == ReferenceStyle.ColumnRelative) 
					{ 
						Relativity = ReferenceStyle.Relative; 
					} 
					else if (_Relativity == ReferenceStyle.Relative) 
					{ 
						Relativity = ReferenceStyle.Absolute; 
					} 
					timer1_Tick(null, null);
				} 
				else 
				{ 
					_Entered = false; 
					timer1.Enabled = false; 
				} 			
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}

		}

		private void refedit_KeyUp(object sender, KeyEventArgs e)
		{
			try 
			{ 
				if(DesignMode) return;

				if(!Visible || !Enabled) return;
			
				string thisKey = ((TextBox)(sender)).Text; 
				if (thisKey.IndexOf("!") != -1) 
				{ 
					thisSheetName = thisKey.Substring(0, thisKey.IndexOf("!")); 
					if (thisSheetName.StartsWith("["))
					{
						thisBookName = thisSheetName.Substring(0, thisSheetName.IndexOf("]") + 1);
						thisSheetName = thisSheetName.Substring(thisBookName.Length);
						thisBookName = thisBookName.Substring(1, thisBookName.Length - 2);
					}
				} 
				else 
				{ 
					thisSheetName = ((Excel.Worksheet)(_Excel.ActiveSheet)).Name; 
				} 
				if (thisSheetName.StartsWith("["))
				{
					thisBookName = thisSheetName.Substring(0, thisSheetName.IndexOf("]") + 1);
					thisSheetName = thisSheetName.Substring(thisBookName.Length);
					thisBookName = thisBookName.Substring(1, thisBookName.Length - 2);
				}
				thisSheetName = thisSheetName.Replace("'", ""); 
				if (_ActiveSheetName != thisSheetName) 
				{ 
					try 
					{ 
						((Excel._Worksheet)(_Excel.Worksheets[thisSheetName])).Activate(); 
					} 
					catch
					{ 
					} 
				} 
				if (thisKey.IndexOf("!") != -1) 
				{ 
					_Excel.get_Range(thisKey.Substring(thisKey.IndexOf("!") + 1), Type.Missing).Select(); 
				} 
				else 
				{
					if (thisKey != "") _Excel.get_Range(thisKey, Type.Missing).Select();
				} 
			} 
			catch (Exception ex)
			{ 
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			} 
		}

		private void button1_Click(object sender, EventArgs e)
		{
			try
			{

				if(DesignMode) return;

				if(!Visible || !Enabled) return;

				textRange.Focus();
				_Entered = true; 
				if (CollapseFormStyle == CollapseStyle.None)  return; 

				var collapse = MDIForm ?? thisForm; 
				collapse.SuspendLayout();
				if (isCollapsed) 
				{
					if (oExcel != null)
					{
						oExcel.MoveAfterReturn = excelMoveCursorOnEnter;
					}
					_showEditField = oldshowEditField;
					restoreForm(collapse);
					timer1.Enabled = false;
					OnRestore(e);
				} 
				else 
				{
					if (oExcel != null)
					{
						oExcel.MoveAfterReturn = false;
					}
					
					parentCaption = thisForm.Text; 
					_showEditField = true;
					collapseForm(collapse);
					timer1.Enabled = true;
				} 
				collapse.ResumeLayout();
				isCollapsed = !(isCollapsed); 
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}
		
		private static void killEnterKey()
		{
			var keyStates = new byte[256];

			GetKeyboardState(keyStates);		  //load the keyboard
			keyStates[ENTER_KEY] = 0;								// turn off the enter key
			SetKeyboardState(keyStates);		  //set the new keyboard state
		}
		
		private void timer1_Tick(object sender, EventArgs e)
		{
			
			try 
			{ 
				if(DesignMode) return;

				if(!Visible || !Enabled) return;
			
				if(_Excel==null)
				{
					return;
				}

				if (isCollapsed)
				{
					if (GetKeyState(ENTER_KEY) == 1)
					{
						killEnterKey();
						button1_Click(null, null);
						//return;
					}
				}
				
				if (_ActiveSheetName == "") 
				{ 
					var mSheet = ((Excel.Worksheet)(_Excel.ActiveSheet)); 
					_ActiveSheetName = mSheet.Name; 
				} 
				if (_Entered) 
				{ 
					_Entered = false; 
					if (textRange.Text.IndexOf("!") != -1) 
					{ 
						thisSheetName = textRange.Text.Substring(0, textRange.Text.IndexOf("!")); 
						if (thisSheetName.StartsWith("["))
						{
							thisBookName = thisSheetName.Substring(0, thisSheetName.IndexOf("]") + 1);
							thisSheetName = thisSheetName.Substring(thisBookName.Length);
							thisBookName = thisBookName.Substring(1, thisBookName.Length - 2);
						}
					} 
					else 
					{ 
						thisSheetName = ((Excel.Worksheet)(_Excel.ActiveSheet)).Name; 
					} 
					thisSheetName = thisSheetName.Replace("'", ""); 
					if (_ActiveSheetName != thisSheetName) 
					{ 
						try 
						{ 
							((Excel._Worksheet)(_Excel.Worksheets[thisSheetName])).Activate(); 
						} 
						catch
						{ 
							MessageBox.Show("Cannot access the worksheet referenced in this cell range reference.", "Cell Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
						} 
					} 
					if (textRange.Text.IndexOf("!") != -1) 
					{ 
						_Excel.get_Range(textRange.Text.Substring(textRange.Text.IndexOf("!") + 1), Type.Missing).Select(); 
					} 
					else if(textRange.Text!="")
					{
						try
						{
							_Excel.get_Range(textRange.Text, Type.Missing).Select();
						}
						catch { }
					} 
				} 
				else 
				{ 
					thisSheetName = ((Excel.Worksheet)(_Excel.ActiveSheet)).Name; 
				} 
				if (_ActiveSheetName != thisSheetName) 
				{ 
					if (_AllowMultipleCells) 
					{ 
						if (thisSheetName.IndexOf(" ") != -1) 
						{ 
							if (textRange.Text != "'" + thisSheetName + "'!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing)) 
							{ 
								textRange.Text = "'" + thisSheetName + "'!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing); 
								if(!textRange.IsDisposed)textRange.SelectAll(); 
							} 
						} 
						else 
						{ 
							if (textRange.Text != thisSheetName + "!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing)) 
							{ 
								textRange.Text = thisSheetName + "!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing); 
								if(!textRange.IsDisposed)textRange.SelectAll(); 
							} 
						} 
					} 
					else 
					{ 
						if (thisSheetName.IndexOf(" ") != -1) 
						{ 
							if (textRange.Text != "'" + thisSheetName + "'!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing)) 
							{ 
								textRange.Text = "'" + thisSheetName + "'!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing); 
							} 
						} 
						else 
						{ 
							if (textRange.Text != thisSheetName + "!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing)) 
							{ 
								textRange.Text = thisSheetName + "!" + ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing); 
							} 
						} 
						if (textRange.Text.IndexOf(":") != -1) 
						{ 
							textRange.Text = textRange.Text.Substring(0, textRange.Text.IndexOf(":")); 
						} 
						if(!textRange.IsDisposed)textRange.SelectAll(); 
					} 
				} 
				else 
				{
					if (_AllowMultipleCells) 
					{ 
						if (textRange.Text != ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing)) 
						{ 
							textRange.Text = ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing); 
							if(!textRange.IsDisposed)textRange.SelectAll(); 
						} 
					}
					else if (_AllowTextEntry)
					{
						try
						{
							if (textRange.Text != ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing))
							{
								textRange.Text = ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing);
								if (textRange.Text.IndexOf(":") != -1)
								{
									textRange.Text = textRange.Text.Substring(0, textRange.Text.IndexOf(":"));
								}
								if (!textRange.IsDisposed) textRange.SelectAll();
							}
						}
						catch { }
					}
					else
					{
						if (textRange.Text != ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing))
						{
							textRange.Text = ((Excel.Range)(_Excel.Selection)).get_AddressLocal(RowAbsolute, ColAbsolute, A1, Type.Missing, Type.Missing);
							if (textRange.Text.IndexOf(":") != -1)
							{
								textRange.Text = textRange.Text.Substring(0, textRange.Text.IndexOf(":"));
							}
							if (!textRange.IsDisposed) textRange.SelectAll();
						}
					} 
				}
				if (_DisplayStyle == DisplayStyle.AddressRangeWithSheet || _DisplayStyle == DisplayStyle.AddressRangeWithSheetBook) 
				{ 
					string wsName = ((Excel.Range)(_Excel.Selection)).Worksheet.Name;
					if(wsName.IndexOf(" ")!=-1)
					{
						wsName = "'" + wsName + "'";
					}
					textRange.Text = wsName + "!" + textRange.Text;
				}
				if (_DisplayStyle == DisplayStyle.AddressRangeWithSheetBook) 
				{ 
					textRange.Text = "[" + ((Excel.Workbook)((Excel.Range)(_Excel.Selection)).Worksheet.Parent).Name + "]" + textRange.Text;
				}
				if (_DisplayStyle == DisplayStyle.SumRangeValues) 
				{ 
					textRange.Tag = textRange.Text; 
					textRange.Text = concatStrings(_Excel.get_Range(textRange.Text, Type.Missing).Cells, false);
				} 
				else if (_DisplayStyle == DisplayStyle.ConcatenateRangeValues) 
				{ 
					textRange.Tag = textRange.Text; 
					textRange.Text = concatStrings(_Excel.get_Range(textRange.Text, Type.Missing).Cells, true);
				} 
			} 
			catch //(Exception ex)
			{ 
				//_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			} 
			//finally 
			//{ 
			//} 
		}

		private void excelEvents_SheetSelectionChange(object Sh)
		{
			if(!Visible || !Enabled) return;
			
			if(ParentForm!=null && ParentForm.ActiveControl!=null)
			{
				if(Name != ParentForm.ActiveControl.Name)
				{
					return;
				}
			}

			textRange.Text = "";
			timer1_Tick(null, null);
		}

		private void docEvents_SelectionChange(Excel.Range Target)
		{
			if(!Visible || !Enabled) return;
			
			timer1_Tick(null, null);
		}

		private void textRange_TextChanged(object sender, EventArgs e)
		{
			OnCellChanged(e);
		}

		#endregion

		#region Private Methods
		private string concatStrings(Excel.Range rng, bool AsStrings)
		{
			var myVal = "";
			try
			{
				var mySep = _DisplayResultSeparator;
				for(var x = 1; x <= rng.Rows.Count; x++)
				{
					for(var y = 1; y <= rng.Columns.Count; y++)
					{
						if(AsStrings)
						{
							myVal += ((Excel.Range)rng.get_Item(x, y)).Value2 + mySep;
						}
						else
						{
							if (IsNumeric(((Excel.Range)rng.get_Item(x, y)).Value2)) 
							{ 
								myVal += (double)((Excel.Range)rng.get_Item(x, y)).Value2; 
							} 
						}
					}
				}
				if(myVal.EndsWith("|"))
				{
					myVal = myVal.Substring(0, myVal.Length - 1);
				}
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
			return myVal;
		}

		private void restoreForm(Form collapse)
		{
			try
			{
				collapse.Size = thisFormSize;
				Location = thisLocation; 

				collapse.ActiveControl.SelectNextControl(this, true, false, true, true);
				
				Parent = parentControl;
				
				if (CollapseFormStyle == CollapseStyle.CollapseFormAndFitCellSelector) 
				{ 
					Dock = DockStyle.None;
					Anchor = AnchorStyles.None;
					Size = thisSize;
					Anchor = anchorControl;
					Dock = dockStyle;
				} 
				if(!_AllowCollapsedResize)
				{
					collapse.FormBorderStyle = borderStyle;
				}
				if (MDIForm != null) 
				{ 
					showMDIControls();
					MDIForm.Menu = MDIFormMenu; 
					MDIForm.TopMost = topMostForm; 
				}
				else
				{
					thisForm.TopMost = topMostForm; 
				}
				showControls();
				if (thisFormMenu != null) 
				{ 
					thisForm.Menu = thisFormMenu; 
				} 
				thisForm.Text = parentCaption; 
				Visible = true; 
//				btnState = BtnState.Normal;
				button1.ImageIndex = 0; 
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}
		private void collapseForm(Form collapse)
		{		
			try
			{
				thisFormSize = collapse.Size;
				thisLocation = Location;
				thisSize = Size;

				parentControl = Parent;
				
				bool visState = Visible;
				
				Parent = ParentForm;

				Visible = visState;

				hideControls();
					
				thisForm.Menu = null;
				if (MDIForm != null) 
				{ 
					hideMDIControls();
					MDIForm.Menu = null;
					MDIForm.TopMost = true;
				}
				else
				{
					thisForm.TopMost = true;
				}
//				btnState = BtnState.Pushed;
				button1.ImageIndex = 1;
				collapse.Text = CollapsedFormCaption;

				//resize the form
				if(!_AllowCollapsedResize)
				{
					borderStyle = collapse.FormBorderStyle;
					if(collapse.FormBorderStyle == FormBorderStyle.SizableToolWindow || collapse.FormBorderStyle == FormBorderStyle.FixedToolWindow)
					{
						collapse.FormBorderStyle = FormBorderStyle.FixedToolWindow;
					}
					else
					{
						collapse.FormBorderStyle = FormBorderStyle.FixedDialog;
					}
				}
				collapse.Height = (Height + (collapse.Height - collapse.ClientSize.Height)) + CollapsedHeightAdjustment;
				if (CollapseFormStyle == CollapseStyle.CollapseFormOnly)
				{
					Location = new Point(Left, 0);
				}
				if (CollapseFormStyle == CollapseStyle.CollapseFormAndFitCellSelector) 
				{ 
					Location = new Point(0, 0);
					dockStyle = Dock;
					anchorControl = Anchor;
					Dock = DockStyle.Fill;
				}
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}

		}
		private int getRealHeight(ref Form collapse)
		{
			var addTop = 0;
			var thisParent = Parent;
			while(thisParent.GetType() != collapse.GetType())
			{
				addTop+=thisParent.Top;
				if(thisParent.Parent==null)
				{
					break;
				}
				thisParent = thisParent.Parent;
			}
			return addTop;
		}
		private static bool IsNumeric(object inValue)
		{
			try
			{
				double.Parse(inValue.ToString());
				return true;
			}
			catch
			{
			}
			return false;
		}

		private void hideControls()
		{
			try
			{
				visibles = new Hashtable();
				parents = new ArrayList();

				//We have to start at this Control, and go up the hierarchy
				//So, get this controls parent
				var thisParent = Parent;
				while(thisParent!=null)
				{
					parents.Add(thisParent.Name);
					hideChildControls(thisParent);
					thisParent = thisParent.Parent;
				}
				
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}

		}

		private void hideChildControls(Control container)
		{
			try
			{
				foreach(Control ctrl in container.Controls) 
				{ 
					visibles.Add(ctrl, ctrl.Visible);
					if(ctrl.Name!=Name && parents.IndexOf(ctrl.Name)==-1)
					{
						ctrl.Visible = false;
					}
				}
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}
		private void hideMDIControls()
		{
			try
			{
				MDIvisibles = new Hashtable();
				foreach (Control ctrl in MDIForm.Controls) 
				{ 
					MDIvisibles.Add(ctrl, ctrl.Visible); 
					if(ctrl!=ParentForm && ctrl!=this && ctrl.Name!=Parent.Name)
					{
						ctrl.Visible = false;
					}
				} 
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}

		private void showControls()
		{
			try
			{
				foreach(DictionaryEntry ctrl in visibles) 
				{ 
					((Control)ctrl.Key).Visible = (bool)ctrl.Value;
				}
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}

		private void showMDIControls()
		{
			try
			{
				foreach(DictionaryEntry ctrl in MDIvisibles) 
				{ 
					((Control)ctrl.Key).Visible = (bool)ctrl.Value;
				}
			}
			catch (Exception ex)
			{
				_xlLogger.LogException(MethodBase.GetCurrentMethod().DeclaringType.Name, ex.ToString(), MethodBase.GetCurrentMethod().Name);
			}
		}

		#endregion


	}
}