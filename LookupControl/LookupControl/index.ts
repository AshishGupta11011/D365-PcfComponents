import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class LookupControl
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  /**
   * Empty constructor.
   */
  private _container: HTMLDivElement;
  private _context: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged: () => void;
  private _contentLabel = "Simple Lookup Control";

  // Containers to store and display raw lookup data for both properties
  private _inputData1: HTMLLabelElement;
  private _inputData2: HTMLLabelElement;

  private _entityType1 = "";
  private _defaultViewId1 = "";
  private _entityType2 = "";
  private _defaultViewId2 = "";

  private _selectedItem1: ComponentFramework.LookupValue;
  private _selectedItem2: ComponentFramework.LookupValue;

  private _updateSelected1 = false;

  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    // Add control initialization code
    this._container = container;
    this._notifyOutputChanged = notifyOutputChanged;
    this._context = context;

    this._entityType1 = context.parameters.controlValue.getTargetEntityType();
    this._entityType2 = context.parameters.controlValue.getTargetEntityType();

    this._defaultViewId1 = context.parameters.controlValue.getViewId();
    this._defaultViewId2 = context.parameters.controlValue.getViewId();

    const contentContainer = document.createElement("div");

    const contentLabel = document.createElement("label");
    contentLabel.textContent = this._contentLabel;
    contentContainer.appendChild(contentLabel);

    contentContainer.append(
      document.createElement("br"),
      document.createElement("br")
    );

    this._inputData1 = document.createElement("label");
    this._inputData1.textContent = "";
    contentContainer.appendChild(this._inputData1);

    contentContainer.append(
      document.createElement("br"),
      document.createElement("br")
    );
    const lookupObjectsButton1 = document.createElement("button");
    lookupObjectsButton1.innerText = "Lookup Objects";
    lookupObjectsButton1.onclick = this.performLookupObjects.bind(
      this,
      this._entityType1,
      this._defaultViewId1,
      (value, update = true) => {
        this._selectedItem1 = value;
        this._updateSelected1 = update;
      }
    );

    contentContainer.append(lookupObjectsButton1);
    contentContainer.append(
      document.createElement("br"),
      document.createElement("br")
    );

    // Add element to display raw entity data
    this._inputData2 = document.createElement("label");
    this._inputData2.textContent = "";
    contentContainer.appendChild(this._inputData2);
    contentContainer.append(
      document.createElement("br"),
      document.createElement("br")
    );
    const lookupObjectsButton2 = document.createElement("button");
    lookupObjectsButton2.innerText = "Lookup Objects";
    lookupObjectsButton2.onclick = this.performLookupObjects.bind(
      this,
      this._entityType2,
      this._defaultViewId2,
      (value) => {
        this._selectedItem2 = value;
        this._updateSelected1 = false;
      }
    );

    contentContainer.append(lookupObjectsButton2);
    contentContainer.append(
      document.createElement("br"),
      document.createElement("br")
    );

    this._container.append(contentContainer);
  }

  public performLookupObjects(
    entityType: string,
    viewId: string,
    setSelected: (
      value: ComponentFramework.LookupValue,
      update?: boolean
    ) => void
  ): void {
    const lookupOtions = {
      defaultEntityType: entityType,
      defaultViewId: viewId,
      allowMultiSelect: false,
      entityTypes: [entityType],
      viewIds: [viewId],
    };
    this._context.utils.lookupObjects(lookupOtions).then(
      (success) => {
        if (success && success.length > 0) {
          const selectedRef = success[0];
          const selectedLookupVal: ComponentFramework.LookupValue = {
            id: selectedRef.id,
            name: selectedRef.name,
            entityType: selectedRef.entityType,
          };
          setSelected(selectedLookupVal);
          this._notifyOutputChanged();
        } else {
          setSelected({} as ComponentFramework.LookupValue);
        }
      },
      (error) => {
        console.log(error);
      }
    );
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Add code to update control view
    const lookupValue1 : ComponentFramework.LookupValue = context.parameters.controlValue.raw[0];
    const lookupValue2 : ComponentFramework.LookupValue= context.parameters.controlValue1.raw[0];
    this._context = context;
    this._inputData1.textContent = `name :${lookupValue1.name} entityType : ${lookupValue1.entityType} id : ${lookupValue1.id}`;
    this._inputData2.textContent = `name :${lookupValue2.name} entityType : ${lookupValue2.entityType} id : ${lookupValue2.id}`;;
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return this._updateSelected1
      ? { controlValue: [this._selectedItem1] }
      : { controlValue1: [this._selectedItem2] };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}
